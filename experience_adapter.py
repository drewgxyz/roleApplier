import anthropic
import json
import re
from pathlib import Path
from typing import Dict
from docx import Document
from job_parser import JobInfo

class ExperienceAdapter:
    def __init__(self, api_key: str, original_cv_path: str = None, context_path: str = None, skills_config: dict = None):
        """Initialize with Claude API and optional original CV for reference"""
        self.client = anthropic.Anthropic(api_key=api_key)
        self.original_cv_data = self._extract_cv_data(original_cv_path) if original_cv_path else None
        self.additional_context = self._load_additional_context(context_path) if context_path else {}
        self.skills_config = skills_config or {}
        
        # Handle both old flat list format and new tiered format
        if isinstance(skills_config.get('my_skills'), dict):
            # New tiered format
            self.my_skills = []
            self.skill_tiers = skills_config['my_skills']
            for tier in ['tier_1_core', 'tier_2_major', 'tier_3_specialist', 'tier_4_tools']:
                self.my_skills.extend(self.skill_tiers.get(tier, []))
        else:
            # Old flat format (backward compatibility)
            self.my_skills = skills_config.get('my_skills', [])
            self.skill_tiers = {'tier_1_core': self.my_skills}
            
        self.blacklisted_skills = skills_config.get('blacklisted_skills', [])
    
    def _extract_cv_data(self, cv_path: str) -> Dict:
        """Extract experience data from original CV document"""
        if not cv_path or not Path(cv_path).exists():
            return {}
            
        try:
            doc = Document(cv_path)
            cv_text = "\n".join([para.text for para in doc.paragraphs])
            
            # Extract the original experience data structure
            return {
                "t_rowe_price": {
                    "original_skills": "Python, Java, AWS, RDS, Terraform",
                    "projects": {
                        "bp1": "Led architecture and development of production-grade Python data migration tool for syncing complex relational data across DEV/STAGE/PROD environments with rollback safety, referential integrity validation, and automated testing",
                        "bp2": "Redesigned legacy application with database performance issues, implementing scalable event-driven architecture on AWS using Lambda, SQS, Python, and Java, achieving 60% performance improvement",
                        "bp3": "Architected disaster recovery strategies for legacy core infrastructure, rebuilding critical Java services in Python to enable seamless DR failover and establishing automated backup systems across US-East-2 region, ensuring high availability for internal finance services",
                        "bp4": "Developed high-performance data loaders for internal directory system, integrating Active Directory data into RDS and OpenSearch to optimize search functionality and reduce response times"
                    }
                },
                "aws": {
                    "original_skills": "Python, AWS, Java, Ruby, Linux System Administration",
                    "projects": {
                        "bp1": "Optimized Region Build deployment process, reducing required time by 40-55% across 15+ Service Catalog services and pipelines",
                        "bp2": "Deployed Service Catalog services across 5 newly launched AWS Regions (UAE, Melbourne, Spain, Zurich, Hyderabad) supporting global expansion",
                        "bp3": "Led security escalation response managing 2,400+ hosts, implementing automated security patching pipelines"
                    }
                },
                "bio": "Proven track record of delivering production-grade solutions that eliminate manual processes, improve system performance by 60%+, and influence organizational security strategy. Key expertise includes architecting Python-based migration tools for complex relational data, implementing scalable event-driven systems on AWS, and optimizing infrastructure deployments across global regions, as well as SQL expertise. Demonstrated ability to present technical findings to leadership and drive strategic decision-making in enterprise environments."
            }
        except Exception as e:
            print(f"Could not extract CV data: {e}")
            return {}
    
    def _load_additional_context(self, context_path: str) -> Dict:
        """Load additional context about projects from JSON file"""
        if not context_path or not Path(context_path).exists():
            return {}
            
        try:
            with open(context_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            print(f"Could not load additional context: {e}")
            return {}
    
    def adapt_experience_to_job(self, job_info: JobInfo) -> Dict:
        """Use AI to adapt experience to match job requirements"""
        
        if not self.original_cv_data:
            raise ValueError("No original CV data loaded. Please provide original CV path.")
        
        # Filter job skills based on my actual skills and blacklist
        relevant_skills = self._filter_job_skills_with_priority(job_info.required_skills + job_info.preferred_skills)
        
        # Assess job relevance to determine content strategy
        relevance_score = self._assess_job_relevance(job_info, relevant_skills)
        content_strategy = self._determine_content_strategy(relevance_score)
        
        # Include skills filtering information in context
        skills_section = f"""
        MY ACTUAL SKILLS (tiered by priority):
        Tier 1 Core: {', '.join(self.skill_tiers.get('tier_1_core', []))}
        Tier 2 Major: {', '.join(self.skill_tiers.get('tier_2_major', []))}
        Tier 3 Specialist: {', '.join(self.skill_tiers.get('tier_3_specialist', []))}
        Tier 4 Tools: {', '.join(self.skill_tiers.get('tier_4_tools', []))}
        
        BLACKLISTED SKILLS (never mention these):
        {', '.join(self.blacklisted_skills)}
        
        JOB RELEVANT SKILLS (prioritized by tier):
        {', '.join(relevant_skills)}
        
        CONTENT STRATEGY FOR THIS JOB:
        Job Relevance Score: {relevance_score}/10
        Content Strategy: {content_strategy['name']}
        Bio Length: {content_strategy['bio_sentences']} sentences
        T. Rowe Price Detail Level: {content_strategy['trp_detail']}
        AWS Detail Level: {content_strategy['aws_detail']}
        Expertise Skills Count: {content_strategy['expertise_count']}
        """
        
        # Include additional context in the prompt
        context_section = ""
        if self.additional_context:
            context_section = f"""
        ADDITIONAL PROJECT CONTEXT (use this to enhance/adapt the bullet points):
        {json.dumps(self.additional_context, indent=2)}
        
        Instructions for using context:
        - Use the additional context to add specific details, technologies, or achievements that match the job requirements
        - If context mentions technologies/skills relevant to the job, incorporate them naturally
        - Use metrics, team sizes, or technical details from context when they strengthen the bullet point
        - Don't use context that's irrelevant to the target job
        """
        
        prompt = f"""
        I need to customize my CV for this specific job opportunity. I want to keep the exact same structure but adapt the content to highlight relevant skills and experiences.

        TARGET JOB:
        - Position: {job_info.job_title} at {job_info.company_name}
        - Required Skills: {', '.join(job_info.required_skills)}
        - Preferred Skills: {', '.join(job_info.preferred_skills)}
        - Key Responsibilities: {', '.join(job_info.key_responsibilities)}
        - Industry: {job_info.industry}

        {skills_section}

        MY CURRENT EXPERIENCE (to adapt from):
        {json.dumps(self.original_cv_data, indent=2)}
        
        {context_section}

        CRITICAL INSTRUCTIONS - ADAPTIVE LENGTH MANAGEMENT:
        1. MUST fit on exactly 1 page - use the content strategy provided
        2. Bio: Use exactly {content_strategy['bio_sentences']} sentences
        3. T. Rowe Price bullets: {content_strategy['trp_detail']} detail level
        4. AWS bullets: {content_strategy['aws_detail']} detail level  
        5. Expertise: Include exactly {content_strategy['expertise_count']} skills
        6. Tech stacks: Maximum 8-10 technologies from MY ACTUAL SKILLS list
        7. NEVER mention any skills from the BLACKLISTED SKILLS list
        8. Prioritize most relevant content first
        9. Include specific metrics where possible but keep within length limits
        10. Use job-relevant keywords naturally

        JOB TITLE RESTRICTIONS (CRITICAL):
        - NEVER call me "Senior" anything in the bio
        - NEVER use job titles higher than my actual level
        - Acceptable titles ONLY: "Software Engineer", "Software Developer", "Mid-level Software Engineer", "Mid-level Software Developer"
        - Even if job posting is for "Senior" roles, stick to my actual level
        - Focus bio on experience and skills, not inflated titles

        DETAIL LEVEL DEFINITIONS:
        - HIGH: 25-35 words, include specific technologies, metrics, and technical details
        - MEDIUM: 18-25 words, include key technologies and one metric/outcome
        - LOW: 12-18 words, focus on impact and one key technology
        - MINIMAL: 8-12 words, essential impact only

        BULLET POINT REQUIREMENTS:
        - ALWAYS mention specific technologies from the tech stacks in bullet points
        - Include deliverable business value and quantified results when space allows
        - Use action verbs and technical specificity
        - Prioritize content relevance to the target job
        - If detail level is HIGH, include multiple technologies per bullet point
        - If detail level is LOW/MINIMAL, focus on most relevant single technology

        Return ONLY a JSON object with these exact fields:
        {{
            "bio": "Bio paragraph with exactly {content_strategy['bio_sentences']} sentences - NEVER use 'Senior' or inflated job titles",
            "expertise": ["Exactly {content_strategy['expertise_count']} most relevant skills from MY ACTUAL SKILLS"],
            "t": {{
                "skills": "Concise tech stack list (only from MY ACTUAL SKILLS)",
                "bp1": "{content_strategy['trp_detail']} detail bullet point mentioning specific technologies and business impact",
                "bp2": "{content_strategy['trp_detail']} detail bullet point with technical specifics and metrics", 
                "bp3": "{content_strategy['trp_detail']} detail bullet point highlighting technologies and value",
                "bp4": "{content_strategy['trp_detail']} detail bullet point with technical detail and outcomes"
            }},
            "a": {{
                "skills": "Concise tech stack list (only from MY ACTUAL SKILLS)",
                "bp1": "{content_strategy['aws_detail']} detail bullet point with efficiency/scale metrics",
                "bp2": "{content_strategy['aws_detail']} detail bullet point emphasizing scope and scale",
                "bp3": "{content_strategy['aws_detail']} detail bullet point with compliance/security impact"
            }}
        }}

        CRITICAL: Follow the exact content strategy provided. Technologies in tech stacks MUST appear in bullet points.
        NEVER inflate job titles - use only "Software Engineer", "Software Developer", or "Mid-level" versions.
        Adjust content density based on detail levels to ensure 1-page fit.
        """

        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=2000,
                messages=[{"role": "user", "content": prompt}]
            )
            
            response_text = response.content[0].text
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            
            if json_match:
                return json.loads(json_match.group())
            else:
                raise ValueError("Could not extract JSON from AI response")
                
        except Exception as e:
            print(f"Error adapting experience: {e}")
            raise
    
    def _filter_job_skills(self, job_skills: list[str]) -> list[str]:
        """Filter job skills to only include ones I actually have"""
        relevant_skills = []
        
        for job_skill in job_skills:
            # Check if this job skill matches any of my skills (case insensitive)
            for my_skill in self.my_skills:
                if (job_skill.lower() in my_skill.lower() or 
                    my_skill.lower() in job_skill.lower() or
                    job_skill.lower() == my_skill.lower()):
                    
                    # Make sure it's not blacklisted
                    is_blacklisted = any(
                        blacklisted.lower() in job_skill.lower() or
                        job_skill.lower() in blacklisted.lower()
                        for blacklisted in self.blacklisted_skills
                    )
                    
                    if not is_blacklisted and my_skill not in relevant_skills:
                        relevant_skills.append(my_skill)
                        break
        
        return relevant_skills
    
    def _filter_job_skills_with_priority(self, job_skills: list[str]) -> list[str]:
        """Filter job skills to only include ones I actually have, prioritized by tier"""
        relevant_skills_by_tier = {
            'tier_1_core': [],
            'tier_2_major': [],
            'tier_3_specialist': [],
            'tier_4_tools': []
        }
        
        for job_skill in job_skills:
            # Check if this job skill matches any of my skills (case insensitive)
            for tier_name, tier_skills in self.skill_tiers.items():
                for my_skill in tier_skills:
                    if (job_skill.lower() in my_skill.lower() or 
                        my_skill.lower() in job_skill.lower() or
                        job_skill.lower() == my_skill.lower()):
                        
                        # Make sure it's not blacklisted
                        is_blacklisted = any(
                            blacklisted.lower() in job_skill.lower() or
                            job_skill.lower() in blacklisted.lower()
                            for blacklisted in self.blacklisted_skills
                        )
                        
                        if not is_blacklisted and my_skill not in relevant_skills_by_tier[tier_name]:
                            relevant_skills_by_tier[tier_name].append(my_skill)
                            break
        
        # Flatten in priority order
        prioritized_skills = []
        for tier in ['tier_1_core', 'tier_2_major', 'tier_3_specialist', 'tier_4_tools']:
            prioritized_skills.extend(relevant_skills_by_tier[tier])
        
        return prioritized_skills
    
    def _assess_job_relevance(self, job_info: JobInfo, relevant_skills: list[str]) -> int:
        """Assess how relevant this job is to my experience (1-10 scale)"""
        relevance_score = 0
        
        # Check for core technology matches
        core_matches = 0
        for skill in self.skill_tiers.get('tier_1_core', []):
            if any(skill.lower() in req.lower() for req in job_info.required_skills + job_info.preferred_skills):
                core_matches += 1
        
        # Score based on core technology alignment
        if core_matches >= 3:
            relevance_score += 4
        elif core_matches >= 2:
            relevance_score += 3
        elif core_matches >= 1:
            relevance_score += 2
        
        # Check for experience domain matches
        finance_keywords = ['finance', 'financial', 'trading', 'banking', 'investment']
        cloud_keywords = ['cloud', 'aws', 'infrastructure', 'devops', 'deployment']
        data_keywords = ['data', 'etl', 'migration', 'database', 'analytics']
        
        job_text = f"{job_info.job_title} {' '.join(job_info.key_responsibilities)} {job_info.industry}".lower()
        
        domain_matches = 0
        if any(keyword in job_text for keyword in finance_keywords):
            domain_matches += 2
        if any(keyword in job_text for keyword in cloud_keywords):
            domain_matches += 2  
        if any(keyword in job_text for keyword in data_keywords):
            domain_matches += 2
            
        relevance_score += min(domain_matches, 4)
        
        # Bonus for seniority level match
        if any(term in job_info.years_experience.lower() for term in ['2-3', '3-5', 'mid', 'senior']):
            relevance_score += 2
            
        return min(relevance_score, 10)
    
    def _determine_content_strategy(self, relevance_score: int) -> dict:
        """Determine content strategy based on job relevance"""
        
        if relevance_score >= 8:
            # High relevance - maximize detail for best match
            return {
                'name': 'HIGH_RELEVANCE',
                'bio_sentences': 3,  # Shorter bio for more bullet space
                'trp_detail': 'HIGH',
                'aws_detail': 'MEDIUM',
                'expertise_count': 12
            }
        elif relevance_score >= 6:
            # Medium-high relevance - balanced approach
            return {
                'name': 'MEDIUM_HIGH_RELEVANCE', 
                'bio_sentences': 3,
                'trp_detail': 'MEDIUM',
                'aws_detail': 'MEDIUM',
                'expertise_count': 10
            }
        elif relevance_score >= 4:
            # Medium relevance - conservative approach
            return {
                'name': 'MEDIUM_RELEVANCE',
                'bio_sentences': 4,
                'trp_detail': 'MEDIUM',
                'aws_detail': 'LOW',
                'expertise_count': 10
            }
        else:
            # Low relevance - minimal approach
            return {
                'name': 'LOW_RELEVANCE',
                'bio_sentences': 4,
                'trp_detail': 'LOW',
                'aws_detail': 'MINIMAL',
                'expertise_count': 8
            }