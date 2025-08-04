import anthropic
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from datetime import datetime
from job_parser import JobInfo

class CoverLetterGenerator:
    def __init__(self, api_key: str):
        """Initialize with Claude API key"""
        self.client = anthropic.Anthropic(api_key=api_key)
    
    def generate_cover_letter(self, job_info: JobInfo, output_path: str):
        """Generate a tailored cover letter for the job"""
        
        # Research company context (basic)
        company_context = self._get_company_context(job_info.company_name, job_info.industry)
        
        prompt = f"""
        Write a professional cover letter for this software engineering position. Keep it genuine, not overly enthusiastic, and focused on technical fit.

        JOB DETAILS:
        - Position: {job_info.job_title}
        - Company: {job_info.company_name}
        - Location: {job_info.location}
        - Industry: {job_info.industry}
        - Required Skills: {', '.join(job_info.required_skills[:8])}
        - Key Responsibilities: {', '.join(job_info.key_responsibilities[:5])}

        MY BACKGROUND:
        - 2.5 years experience as Software Engineer
        - Currently at T. Rowe Price (financial services)
        - Previously at AWS (cloud infrastructure)
        - Key skills: Python, Java, AWS, SQL, data migration, cloud architecture
        - Education: BSc Cyber Security from Warwick University (2022)
        - Location: London, UK

        COMPANY CONTEXT:
        {company_context}

        INSTRUCTIONS:
        1. Keep it professional but not overly formal
        2. 2-3 paragraphs maximum (shorter for 1-page fit)
        3. Mention 2-3 most relevant technical skills that match the job
        4. Reference company/industry in a natural way (not forced enthusiasm)
        5. Focus on what I can contribute, not what I want to gain
        6. Keep it genuine - avoid AI-sounding phrases
        7. Don't oversell or exaggerate - I'm mid-level, not senior
        8. Be concise - target 75% of page maximum
        9. End with simple professional closing
        10. Do NOT include placeholder text like [Your name] - use Drew Gillies throughout
        11. Do NOT repeat closing phrases or signatures

        Return ONLY the cover letter text without any placeholders, signatures, or repeated closings.
        """

        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=600,  # Reduced for shorter letters
                messages=[{"role": "user", "content": prompt}]
            )
            
            cover_letter_text = response.content[0].text.strip()
            
            # Generate PDF
            self._create_cover_letter_pdf(cover_letter_text, job_info, output_path)
            
            return cover_letter_text
            
        except Exception as e:
            print(f"Error generating cover letter: {e}")
            raise
    
    def _get_company_context(self, company_name: str, industry: str) -> str:
        """Get basic company context without making additional API calls"""
        # Simple company context based on known patterns
        context_map = {
            'fintech': 'financial technology solutions',
            'finance': 'financial services',
            'banking': 'banking and financial services',
            'trading': 'trading and investment',
            'healthcare': 'healthcare technology',
            'e-commerce': 'online retail and e-commerce',
            'saas': 'software-as-a-service solutions',
            'startup': 'innovative technology solutions',
            'consulting': 'technology consulting services',
            'media': 'digital media and entertainment',
            'gaming': 'gaming and interactive entertainment',
            'education': 'educational technology',
            'logistics': 'logistics and supply chain technology'
        }
        
        # Try to match industry or company keywords
        company_lower = company_name.lower()
        industry_lower = industry.lower()
        
        for keyword, description in context_map.items():
            if keyword in industry_lower or keyword in company_lower:
                return f"As a company focused on {description}, {company_name} likely values technical excellence and scalable solutions."
        
        return f"{company_name} operates in the {industry} sector, requiring robust and reliable software solutions."
    
    def _create_cover_letter_pdf(self, cover_letter_text: str, job_info: JobInfo, output_path: str):
        """Create a professional PDF cover letter"""
        
        doc = SimpleDocTemplate(output_path, pagesize=letter,
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=18)
        
        # Define styles
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
        
        # Build the PDF content
        story = []
        
        # Header with contact info
        header_text = """
        <b>Drew Gillies</b><br/>
        Software Engineer<br/>
        London, UK<br/>
        drew.gillies@hotmail.co.uk<br/>
        07950 298726<br/>
        linkedin.com/in/drew-gillies
        """
        story.append(Paragraph(header_text, header_style))
        story.append(Spacer(1, 20))
        
        # Date
        date_text = datetime.now().strftime("%B %d, %Y")
        story.append(Paragraph(date_text, normal_style))
        story.append(Spacer(1, 12))
        
        # Hiring manager address
        if job_info.company_name and job_info.company_name != "Unknown Company":
            address_text = f"""
            Hiring Manager<br/>
            {job_info.company_name}<br/>
            {job_info.location if job_info.location != "Not specified" else ""}
            """
            story.append(Paragraph(address_text, normal_style))
            story.append(Spacer(1, 12))
        
        # Subject line
        subject_text = f"<b>Re: {job_info.job_title} Position</b>"
        story.append(Paragraph(subject_text, normal_style))
        story.append(Spacer(1, 12))
        
        # Cover letter body
        paragraphs = cover_letter_text.split('\n\n')
        for paragraph in paragraphs:
            if paragraph.strip():
                # Clean up any template artifacts
                clean_paragraph = paragraph.strip()
                clean_paragraph = clean_paragraph.replace('[Your name]', 'Drew Gillies')
                clean_paragraph = clean_paragraph.replace('[Your Name]', 'Drew Gillies')
                clean_paragraph = clean_paragraph.replace('Best regards, [Your name]', '')
                clean_paragraph = clean_paragraph.replace('Best regards,', '')
                clean_paragraph = clean_paragraph.replace('Sincerely, [Your name]', '')
                
                if clean_paragraph:  # Only add non-empty paragraphs
                    story.append(Paragraph(clean_paragraph, normal_style))
                    story.append(Spacer(1, 12))
        
        # Professional closing
        closing_text = """
        Sincerely,<br/>
        <br/>
        <b>Drew Gillies</b>
        """
        story.append(Spacer(1, 12))
        story.append(Paragraph(closing_text, normal_style))
        
        # Build PDF
        doc.build(story)