"""
Flask Web Application for Batch CV and Cover Letter Generation
Uses Claude 3.5 Sonnet for ATS-friendly CV generation
With configurable skills, experiences, and settings stored in SQLite
"""

import os
import json
import re
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from dotenv import load_dotenv
import anthropic
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from docx import Document
import zipfile
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse

# Import database functions
from database import (
    get_all_skills, get_approved_skills, get_blacklisted_skills,
    get_skills_by_tier, get_experiences, get_settings, update_setting,
    add_skill, update_skill, delete_skill, add_experience, update_experience,
    add_bullet_point, update_bullet_point, delete_bullet_point, get_db
)

# Load environment variables
load_dotenv()

app = Flask(__name__)

# Configuration
ANTHROPIC_API_KEY = os.getenv('ANTHROPIC_API_KEY')
TEMPLATE_PATH = 'templates/template copy.docx'

# Claude 3.5 Sonnet - Best for ATS-friendly CVs
# Reasons: Excellent at following precise formatting, consistent structure,
# professional language, and understanding context for skill matching
CLAUDE_MODEL = "claude-3-5-sonnet-20241022"


class BatchCVGenerator:
    """Handles batch generation of CVs and cover letters using Claude 3.5 Sonnet"""
    
    def __init__(self, api_key: str):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.skill_tiers = get_skills_by_tier()
        self.approved_skills = get_approved_skills()
        self.blacklisted_skills = get_blacklisted_skills()
        self.experiences = get_experiences()
        self.settings = get_settings()
    
    def parse_job_description(self, job_text: str) -> dict:
        """Parse job description using Claude"""
        prompt = f"""
        Analyze this job description and extract key information in JSON format.
        
        Job Description:
        {job_text}

        Return ONLY a valid JSON object with these fields:
        - job_title: The job title/position
        - company_name: Company name (if mentioned, otherwise "Unknown Company")
        - location: Location/city
        - required_skills: Array of technical skills mentioned as required
        - preferred_skills: Array of nice-to-have skills
        - years_experience: Experience level required
        - key_responsibilities: Array of main job responsibilities (3-5)
        - industry: Industry sector
        - remote_policy: "Remote", "Hybrid", "On-site", or "Not specified"
        """
        
        response = self.client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}]
        )
        
        response_text = response.content[0].text
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        raise ValueError("Could not parse job description")
    
    def _filter_and_match_skills(self, job_skills: list) -> list:
        """
        Filter job skills to only include approved ones, prioritized by tier.
        This is the smart matching - if job mentions Kafka/SQS/etc and we have it approved, include it.
        """
        matched_skills = {
            'tier_1_core': [],
            'tier_2_major': [],
            'tier_3_specialist': [],
            'tier_4_tools': [],
        }
        
        for job_skill in job_skills:
            job_skill_lower = job_skill.lower().strip()
            
            # Check against each tier
            for tier, tier_skills in self.skill_tiers.items():
                for my_skill in tier_skills:
                    my_skill_lower = my_skill.lower()
                    
                    # Flexible matching
                    if (job_skill_lower in my_skill_lower or 
                        my_skill_lower in job_skill_lower or
                        job_skill_lower == my_skill_lower):
                        
                        # Check not blacklisted
                        if my_skill not in self.blacklisted_skills:
                            if my_skill not in matched_skills[tier]:
                                matched_skills[tier].append(my_skill)
                        break
        
        # Flatten in priority order
        result = []
        for tier in ['tier_1_core', 'tier_2_major', 'tier_3_specialist', 'tier_4_tools']:
            result.extend(matched_skills[tier])
        
        return result
    
    def _build_experience_context(self, matched_skills: list) -> str:
        """Build experience context with mutable tech stacks based on matched skills"""
        context_parts = []
        
        for exp in self.experiences:
            exp_context = f"\n{exp['company']} - {exp['role']}"
            if exp['start_date']:
                exp_context += f" ({exp['start_date']} - {'Present' if exp['is_current'] else exp['end_date']})"
            
            bullets = []
            for bp in exp.get('bullet_points', []):
                desc = bp['base_description']
                
                # If there are tech placeholders, try to fill them with matched skills
                if bp.get('tech_placeholder'):
                    placeholders = bp['tech_placeholder'].split(',')
                    for ph in placeholders:
                        ph = ph.strip()
                        # Find a relevant skill from matched skills
                        replacement = self._find_skill_for_placeholder(ph, matched_skills)
                        if replacement:
                            desc = desc.replace('{' + ph + '}', replacement)
                        else:
                            # Use a generic term if no match
                            desc = desc.replace('{' + ph + '}', ph.title())
                
                bullets.append(f"  - {desc}")
            
            exp_context += "\n" + "\n".join(bullets)
            context_parts.append(exp_context)
        
        return "\n".join(context_parts)
    
    def _find_skill_for_placeholder(self, placeholder: str, matched_skills: list) -> str:
        """Find a matching skill for a placeholder based on context"""
        placeholder_lower = placeholder.lower()
        
        # Mapping of placeholder types to skill categories
        mappings = {
            'tech': ['Python', 'Java', 'JavaScript'],
            'lang1': ['Java', 'Python'],
            'lang2': ['Python', 'Java'],
            'cloud': ['AWS', 'Lambda', 'EC2'],
            'services': ['Lambda', 'SQS', 'S3', 'RDS'],
            'database': ['RDS', 'PostgreSQL', 'MySQL', 'DynamoDB'],
            'search': ['OpenSearch', 'Elasticsearch'],
            'messaging': ['SQS', 'Kafka', 'RabbitMQ'],
        }
        
        # Check if placeholder matches a mapping
        if placeholder_lower in mappings:
            for skill in mappings[placeholder_lower]:
                if skill in matched_skills:
                    return skill
        
        # Otherwise try direct match
        for skill in matched_skills:
            if placeholder_lower in skill.lower():
                return skill
        
        return None
    
    def generate_cv_and_cover_letter(self, job_info: dict) -> dict:
        """Generate both CV content and cover letter in a single LLM call"""
        
        # Get matched skills from job requirements
        all_job_skills = job_info.get('required_skills', []) + job_info.get('preferred_skills', [])
        matched_skills = self._filter_and_match_skills(all_job_skills)
        
        # Build experience context with matched tech
        experience_context = self._build_experience_context(matched_skills)
        
        # Get settings
        settings = self.settings
        expertise_count = int(settings.get('expertise_count', '14'))
        
        skills_section = f"""
        MY APPROVED SKILLS (by priority tier):
        Tier 1 Core: {', '.join(self.skill_tiers.get('tier_1_core', []))}
        Tier 2 Major: {', '.join(self.skill_tiers.get('tier_2_major', []))}
        Tier 3 Specialist: {', '.join(self.skill_tiers.get('tier_3_specialist', []))}
        Tier 4 Tools: {', '.join(self.skill_tiers.get('tier_4_tools', []))}
        
        BLACKLISTED SKILLS (never mention these):
        {', '.join(self.blacklisted_skills[:30])}... (and more)
        
        MATCHED SKILLS FOR THIS JOB (prioritized):
        {', '.join(matched_skills)}
        """
        
        # Sanitize job title
        forbidden_keywords = r'\b(Senior|Lead|Principal|Staff|I|II|III|IV|DevOps)\b'
        sanitized_title = re.sub(forbidden_keywords, '', job_info.get('job_title', ''), flags=re.IGNORECASE).strip()
        
        prompt = f"""
        Generate BOTH a customized CV and cover letter for this job application in a SINGLE response.
        
        CRITICAL: This CV must be ATS-friendly. Use clear section headers, standard formatting, and include exact keyword matches from the job description.

        TARGET JOB:
        - Position: {sanitized_title} at {job_info.get('company_name', 'Unknown Company')}
        - Location: {job_info.get('location', 'Not specified')}
        - Industry: {job_info.get('industry', 'Technology')}
        - Required Skills: {', '.join(job_info.get('required_skills', []))}
        - Preferred Skills: {', '.join(job_info.get('preferred_skills', []))}
        - Key Responsibilities: {', '.join(job_info.get('key_responsibilities', []))}

        {skills_section}

        MY EXPERIENCE (adapt and enhance these):
        {experience_context}

        MY BACKGROUND FOR COVER LETTER:
        - {settings.get('years_experience', '2.5')} years experience as Software Engineer
        - Currently at T. Rowe Price (financial services)
        - Previously at AWS (cloud infrastructure)
        - Education: {settings.get('education', 'BSc Cyber Security from Warwick University (2022)')}
        - Location: {settings.get('user_location', 'London, UK')}

        CRITICAL INSTRUCTIONS:
        1. CV must fit on exactly 1 page - be concise but impactful
        2. Bio: 2-3 sentences. NEVER use "Senior" or inflated titles. Use: Software Engineer, Software Developer, Backend Engineer
        3. Bullet points: 50-70 words each, include specific technologies from MATCHED SKILLS and quantified metrics
        4. Expertise: Exactly {expertise_count} skills from MATCHED SKILLS list, programming languages first
        5. Tech stacks: Use technologies from MATCHED SKILLS that appear in the job description
        6. NEVER mention any BLACKLISTED SKILLS
        7. For ATS: Use exact keyword matches from job description where they match our approved skills
        8. Cover letter: 2-3 paragraphs, professional, mention 2-3 relevant matched skills

        Return ONLY a JSON object with this exact structure:
        {{
            "cv": {{
                "bio": "Updated bio paragraph - ATS optimized with keywords",
                "expertise": ["List of exactly {expertise_count} skills from MATCHED SKILLS"],
                "t": {{
                    "skills": "Comma-separated tech stack string using MATCHED SKILLS",
                    "bp1": "First bullet point with specific tech and metrics",
                    "bp2": "Second bullet point with specific tech and metrics",
                    "bp3": "Third bullet point with specific tech and metrics",
                    "bp4": "Fourth bullet point with specific tech and metrics"
                }},
                "a": {{
                    "skills": "Comma-separated tech stack string using MATCHED SKILLS",
                    "bp1": "First bullet point",
                    "bp2": "Second bullet point",
                    "bp3": "Third bullet point"
                }}
            }},
            "cover_letter": "Complete cover letter text (2-3 paragraphs). Use name: {settings.get('user_name', 'Drew Gillies')}. No placeholders."
        }}
        """
        
        response = self.client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=3500,
            messages=[{"role": "user", "content": prompt}]
        )
        
        response_text = response.content[0].text
        # Clean up potential issues
        response_text = re.sub(r',\s*([\}\]])', r'\1', response_text)
        
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            return json.loads(json_match.group())
        raise ValueError("Could not parse CV generation response")
    
    def create_cover_letter_pdf(self, cover_letter_text: str, job_info: dict, output_path: str):
        """Create a professional PDF cover letter"""
        settings = self.settings
        
        doc = SimpleDocTemplate(output_path, pagesize=letter,
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=18)
        
        styles = getSampleStyleSheet()
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=styles['Normal'],
            fontSize=12,
            fontName='Helvetica-Bold',
            spaceAfter=12,
        )
        
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=11,
            fontName='Helvetica',
            spaceAfter=12,
            leading=14,
        )
        
        story = []
        
        # Header
        header_text = f"""
        <b>{settings.get('user_name', 'Drew Gillies')}</b><br/>
        Software Engineer<br/>
        {settings.get('user_location', 'London, UK')}<br/>
        {settings.get('user_email', 'drew.gillies@hotmail.co.uk')}<br/>
        {settings.get('user_phone', '07950 298726')}<br/>
        {settings.get('user_linkedin', 'linkedin.com/in/drew-gillies')}
        """
        story.append(Paragraph(header_text, header_style))
        story.append(Spacer(1, 20))
        
        # Date
        date_text = datetime.now().strftime("%B %d, %Y")
        story.append(Paragraph(date_text, normal_style))
        story.append(Spacer(1, 12))
        
        # Company address
        company_name = job_info.get('company_name', 'Unknown Company')
        location = job_info.get('location', '')
        if company_name != "Unknown Company":
            address_text = f"""
            Hiring Manager<br/>
            {company_name}<br/>
            {location if location != "Not specified" else ""}
            """
            story.append(Paragraph(address_text, normal_style))
            story.append(Spacer(1, 12))
        
        # Subject
        job_title = job_info.get('job_title', 'Software Engineer')
        subject_text = f"<b>Re: {job_title} Position</b>"
        story.append(Paragraph(subject_text, normal_style))
        story.append(Spacer(1, 12))
        
        # Body
        paragraphs = cover_letter_text.split('\n\n')
        for paragraph in paragraphs:
            if paragraph.strip():
                clean_paragraph = paragraph.strip()
                clean_paragraph = clean_paragraph.replace('[Your name]', settings.get('user_name', 'Drew Gillies'))
                clean_paragraph = clean_paragraph.replace('[Your Name]', settings.get('user_name', 'Drew Gillies'))
                clean_paragraph = clean_paragraph.replace('Best regards,', '')
                clean_paragraph = clean_paragraph.replace('Sincerely,', '')
                
                if clean_paragraph:
                    story.append(Paragraph(clean_paragraph, normal_style))
                    story.append(Spacer(1, 12))
        
        story.append(Spacer(1, 12))
        doc.build(story)
    
    def create_cv_docx(self, cv_data: dict, job_info: dict, output_path: str):
        """Create CV document from template"""
        if not Path(TEMPLATE_PATH).exists():
            raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
        
        doc = Document(TEMPLATE_PATH)
        
        # Prepare replacements
        replacements = {
            'bio': cv_data.get('bio', ''),
            'expertise': cv_data.get('expertise', []),
            't.skills': cv_data.get('t', {}).get('skills', ''),
            't.bp1': cv_data.get('t', {}).get('bp1', ''),
            't.bp2': cv_data.get('t', {}).get('bp2', ''),
            't.bp3': cv_data.get('t', {}).get('bp3', ''),
            't.bp4': cv_data.get('t', {}).get('bp4', ''),
            'a.skills': cv_data.get('a', {}).get('skills', ''),
            'a.bp1': cv_data.get('a', {}).get('bp1', ''),
            'a.bp2': cv_data.get('a', {}).get('bp2', ''),
            'a.bp3': cv_data.get('a', {}).get('bp3', ''),
        }
        
        # Split expertise for two columns
        expertise = cv_data.get('expertise', [])
        midpoint = (len(expertise) + 1) // 2
        replacements['expertise'] = expertise[:midpoint]
        replacements['expertise2'] = expertise[midpoint:]
        
        # Replace placeholders
        self._replace_placeholders(doc, replacements)
        
        doc.save(output_path)
    
    def _replace_placeholders(self, doc, replacements):
        """Replace placeholders in document"""
        for paragraph in doc.paragraphs:
            self._replace_in_paragraph(paragraph, replacements)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, replacements)
        
        for section in doc.sections:
            for paragraph in section.header.paragraphs:
                self._replace_in_paragraph(paragraph, replacements)
            for paragraph in section.footer.paragraphs:
                self._replace_in_paragraph(paragraph, replacements)
    
    def _replace_in_paragraph(self, paragraph, replacements):
        """Replace placeholders in a paragraph"""
        full_text = paragraph.text
        
        for key, value in replacements.items():
            placeholder = f"{{{{{key}}}}}"
            
            if placeholder in full_text:
                if key in ['t.skills', 'a.skills']:
                    if isinstance(value, list):
                        value_text = ', '.join(value)
                    else:
                        value_text = str(value)
                elif isinstance(value, list):
                    value_text = '\n• '.join(value)
                    value_text = '• ' + value_text if value else ''
                else:
                    value_text = str(value).replace('\n', ' ').strip()
                
                for run in paragraph.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, value_text)


# URL Scraping functions

def scrape_job_url(url: str) -> str:
    """Scrape job description from a URL"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
        }
        
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Remove script and style elements
        for script in soup(["script", "style", "nav", "header", "footer", "aside"]):
            script.decompose()
        
        # Try to find job-specific content containers
        job_content = None
        selectors = [
            '[class*="job-description"]',
            '[class*="jobDescription"]',
            '[class*="job_description"]',
            '[class*="job-details"]',
            '[class*="jobDetails"]',
            '[class*="posting-content"]',
            '[class*="description"]',
            '[id*="job-description"]',
            '[id*="jobDescription"]',
            'article',
            'main',
            '[role="main"]',
        ]
        
        for selector in selectors:
            elements = soup.select(selector)
            if elements:
                best = max(elements, key=lambda x: len(x.get_text()))
                if len(best.get_text().strip()) > 200:
                    job_content = best
                    break
        
        if job_content:
            text = job_content.get_text(separator='\n', strip=True)
        else:
            text = soup.get_text(separator='\n', strip=True)
        
        # Clean up
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        text = '\n'.join(lines)
        
        if len(text) > 8000:
            text = text[:8000] + '\n[Content truncated...]'
        
        return text
        
    except requests.exceptions.Timeout:
        raise Exception(f"Timeout while fetching {url}")
    except requests.exceptions.RequestException as e:
        raise Exception(f"Failed to fetch {url}: {str(e)}")
    except Exception as e:
        raise Exception(f"Error scraping {url}: {str(e)}")


def is_url(text: str) -> bool:
    """Check if text is a valid URL"""
    try:
        result = urlparse(text.strip())
        return all([result.scheme in ['http', 'https'], result.netloc])
    except:
        return False


def parse_urls(input_text: str) -> list:
    """Parse URLs from input text"""
    urls = []
    lines = input_text.strip().split('\n')
    
    for line in lines:
        parts = line.split(',')
        for part in parts:
            url = part.strip()
            if url and is_url(url):
                urls.append(url)
    
    return urls


def parse_job_descriptions(input_text: str) -> list:
    """Parse multiple job descriptions from input text"""
    separators = ['---', '===', '***', '###']
    for sep in separators:
        if sep in input_text:
            parts = input_text.split(sep)
            jobs = [part.strip() for part in parts if part.strip()]
            if len(jobs) > 1:
                return jobs
    
    if '\n\n\n' in input_text:
        parts = input_text.split('\n\n\n')
        jobs = [part.strip() for part in parts if part.strip()]
        if len(jobs) > 1:
            return jobs
    
    return [input_text.strip()] if input_text.strip() else []


# Flask Routes

@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')


@app.route('/settings')
def settings_page():
    """Render the settings page"""
    settings = get_settings()
    skills = get_all_skills(include_blacklisted=True)
    experiences = get_experiences()
    return render_template('settings.html', settings=settings, skills=skills, experiences=experiences)


@app.route('/api/settings', methods=['GET', 'POST'])
def api_settings():
    """Get or update settings"""
    if request.method == 'GET':
        return jsonify(get_settings())
    
    data = request.json
    for key, value in data.items():
        update_setting(key, value)
    
    return jsonify({'success': True})


@app.route('/api/skills', methods=['GET', 'POST'])
def api_skills():
    """Get all skills or add a new skill"""
    if request.method == 'GET':
        return jsonify(get_all_skills(include_blacklisted=True))
    
    data = request.json
    skill_id = add_skill(
        name=data['name'],
        tier=data.get('tier', 'tier_4_tools'),
        category=data.get('category'),
        is_blacklisted=data.get('is_blacklisted', False)
    )
    
    if skill_id:
        return jsonify({'success': True, 'id': skill_id})
    return jsonify({'success': False, 'error': 'Skill already exists'}), 400


@app.route('/api/skills/<int:skill_id>', methods=['PUT', 'DELETE'])
def api_skill(skill_id):
    """Update or delete a skill"""
    if request.method == 'DELETE':
        delete_skill(skill_id)
        return jsonify({'success': True})
    
    data = request.json
    update_skill(
        skill_id,
        name=data.get('name'),
        tier=data.get('tier'),
        is_blacklisted=data.get('is_blacklisted')
    )
    return jsonify({'success': True})


@app.route('/api/experiences', methods=['GET', 'POST'])
def api_experiences():
    """Get all experiences or add a new one"""
    if request.method == 'GET':
        return jsonify(get_experiences())
    
    data = request.json
    exp_id = add_experience(
        company=data['company'],
        role=data['role'],
        start_date=data.get('start_date'),
        end_date=data.get('end_date'),
        is_current=data.get('is_current', False)
    )
    
    return jsonify({'success': True, 'id': exp_id})


@app.route('/api/experiences/<int:exp_id>', methods=['PUT', 'DELETE'])
def api_experience(exp_id):
    """Update or delete an experience"""
    if request.method == 'DELETE':
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM experiences WHERE id = ?', (exp_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    
    data = request.json
    update_experience(exp_id, **data)
    return jsonify({'success': True})


@app.route('/api/bullet_points', methods=['POST'])
def api_add_bullet():
    """Add a bullet point"""
    data = request.json
    bp_id = add_bullet_point(
        experience_id=data['experience_id'],
        base_description=data['base_description'],
        tech_placeholder=data.get('tech_placeholder')
    )
    return jsonify({'success': True, 'id': bp_id})


@app.route('/api/bullet_points/<int:bp_id>', methods=['PUT', 'DELETE'])
def api_bullet(bp_id):
    """Update or delete a bullet point"""
    if request.method == 'DELETE':
        delete_bullet_point(bp_id)
        return jsonify({'success': True})
    
    data = request.json
    update_bullet_point(
        bp_id,
        base_description=data.get('base_description'),
        tech_placeholder=data.get('tech_placeholder')
    )
    return jsonify({'success': True})


@app.route('/generate', methods=['POST'])
def generate():
    """Generate CVs and cover letters for multiple job descriptions"""
    if not ANTHROPIC_API_KEY or ANTHROPIC_API_KEY == 'your_anthropic_api_key_here':
        return jsonify({
            'success': False,
            'error': 'Anthropic API key not configured. Please set ANTHROPIC_API_KEY in .env file.'
        }), 400
    
    data = request.json
    input_text = data.get('job_descriptions', '')
    input_mode = data.get('input_mode', 'text')
    
    if not input_text:
        return jsonify({
            'success': False,
            'error': 'No input provided'
        }), 400
    
    # Parse based on input mode
    jobs = []
    scrape_errors = []
    
    if input_mode == 'urls':
        urls = parse_urls(input_text)
        
        if not urls:
            return jsonify({
                'success': False,
                'error': 'No valid URLs found. Please enter URLs starting with http:// or https://'
            }), 400
        
        for url in urls:
            try:
                job_text = scrape_job_url(url)
                jobs.append({'text': job_text, 'source_url': url})
            except Exception as e:
                scrape_errors.append({'url': url, 'error': str(e)})
    else:
        parsed = parse_job_descriptions(input_text)
        jobs = [{'text': job, 'source_url': None} for job in parsed]
    
    if not jobs:
        error_msg = 'Could not parse any job descriptions'
        if scrape_errors:
            error_msg += '. Scraping errors: ' + '; '.join([f"{e['url']}: {e['error']}" for e in scrape_errors])
        return jsonify({
            'success': False,
            'error': error_msg
        }), 400
    
    generator = BatchCVGenerator(ANTHROPIC_API_KEY)
    results = []
    settings = get_settings()
    
    # Create output directory
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path(f"outputs/batch_{timestamp}")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    for i, job_data in enumerate(jobs):
        job_text = job_data['text']
        source_url = job_data.get('source_url')
        
        try:
            # Parse job description
            job_info = generator.parse_job_description(job_text)
            
            # Generate CV and cover letter in single call
            generated = generator.generate_cv_and_cover_letter(job_info)
            
            # Create safe filename
            company_safe = re.sub(r'[^\w\s-]', '', job_info.get('company_name', 'Company')).replace(' ', '_')
            title_safe = re.sub(r'[^\w\s-]', '', job_info.get('job_title', 'Position')).replace(' ', '_')
            
            # Create job-specific folder
            job_folder = output_dir / f"{i+1}_{company_safe}_{title_safe}"
            job_folder.mkdir(exist_ok=True)
            
            # Save CV
            user_name = settings.get('user_name', 'Drew_Gillies').replace(' ', '_')
            cv_path = job_folder / f"{user_name}_Software_Resume.docx"
            generator.create_cv_docx(generated.get('cv', {}), job_info, str(cv_path))
            
            # Save cover letter
            cover_letter_path = job_folder / f"{user_name}_Cover_Letter_{company_safe}.pdf"
            generator.create_cover_letter_pdf(
                generated.get('cover_letter', ''),
                job_info,
                str(cover_letter_path)
            )
            
            # Save original job description
            job_desc_path = job_folder / "Original_Job_Description.txt"
            with open(job_desc_path, 'w', encoding='utf-8') as f:
                f.write(f"Job Title: {job_info.get('job_title', 'Unknown')}\n")
                f.write(f"Company: {job_info.get('company_name', 'Unknown')}\n")
                f.write(f"Date Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                if source_url:
                    f.write(f"Source URL: {source_url}\n")
                f.write(f"\n{'='*50}\n")
                f.write("ORIGINAL JOB DESCRIPTION:\n")
                f.write(f"{'='*50}\n\n")
                f.write(job_text)
            
            results.append({
                'success': True,
                'job_title': job_info.get('job_title', 'Unknown'),
                'company': job_info.get('company_name', 'Unknown'),
                'source_url': source_url,
                'cv_path': str(cv_path),
                'cover_letter_path': str(cover_letter_path),
                'folder': str(job_folder)
            })
            
        except Exception as e:
            results.append({
                'success': False,
                'error': str(e),
                'source_url': source_url,
                'job_text_preview': job_text[:100] + '...' if len(job_text) > 100 else job_text
            })
    
    # Create zip file of all outputs
    zip_path = output_dir / "all_applications.zip"
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for result in results:
            if result.get('success'):
                folder = Path(result['folder'])
                for file in folder.iterdir():
                    zipf.write(file, f"{folder.name}/{file.name}")
    
    return jsonify({
        'success': True,
        'results': results,
        'output_directory': str(output_dir),
        'zip_file': str(zip_path),
        'total_jobs': len(jobs),
        'successful': sum(1 for r in results if r.get('success')),
        'failed': sum(1 for r in results if not r.get('success')),
        'scrape_errors': scrape_errors
    })


@app.route('/download/<path:filepath>')
def download(filepath):
    """Download a generated file"""
    try:
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 404


if __name__ == '__main__':
    Path('templates').mkdir(exist_ok=True)
    
    print("🚀 Starting CV Generator Web App")
    print("=" * 40)
    print(f"📊 Using Claude Model: {CLAUDE_MODEL}")
    print("   (Best for ATS-friendly CVs)")
    
    if not ANTHROPIC_API_KEY or ANTHROPIC_API_KEY == 'your_anthropic_api_key_here':
        print("⚠️  Warning: ANTHROPIC_API_KEY not set in .env file")
        print("   Create a .env file with: ANTHROPIC_API_KEY=your_key_here")
    else:
        print("✓ Anthropic API key loaded")
    
    print("\n📝 Open http://localhost:5001 in your browser")
    print("⚙️  Settings page: http://localhost:5001/settings")
    app.run(debug=True, port=5001)
