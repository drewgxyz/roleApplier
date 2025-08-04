import anthropic
import json
import re
from pathlib import Path
from typing import Dict
from docx import Document
from job_parser import JobInfo

class ExperienceAdapter:
    def __init__(self, api_key: str, original_cv_path: str = None, context_path: str = None):
        """Initialize with Claude API and optional original CV for reference"""
        self.client = anthropic.Anthropic(api_key=api_key)
        self.original_cv_data = self._extract_cv_data(original_cv_path) if original_cv_path else None
        self.additional_context = self._load_additional_context(context_path) if context_path else {}
    
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

        MY CURRENT EXPERIENCE (to adapt from):
        {json.dumps(self.original_cv_data, indent=2)}
        
        {context_section}

        CRITICAL INSTRUCTIONS:
        1. MUST fit on exactly 1 page - be extremely concise
        2. Bio: Maximum 4 lines, focus on most relevant experience
        3. Each bullet point: Maximum 1-2 lines (about 15-20 words max)
        4. Tech stacks: Comma-separated list, maximum 8-10 technologies
        5. Keep the same professional tone as the original
        6. Include specific metrics where possible but keep brief
        7. Use job-relevant keywords naturally but don't stuff
        8. Prioritize impact and relevance over detail

        FORMATTING REQUIREMENTS:
        - Bio should be 3-4 sentences maximum
        - Each bullet point should be ONE line if possible
        - Focus on action verbs and quantified results
        - Remove unnecessary words and filler

        Return ONLY a JSON object with these exact fields:
        {{
            "bio": "Updated bio paragraph (3-4 sentences max)",
            "t": {{
                "skills": "Concise tech stack list",
                "bp1": "One-line bullet point emphasizing relevant achievement",
                "bp2": "One-line bullet point with performance metric", 
                "bp3": "One-line bullet point highlighting relevant tech",
                "bp4": "One-line bullet point with business impact"
            }},
            "a": {{
                "skills": "Concise tech stack list",
                "bp1": "One-line bullet point with efficiency improvement",
                "bp2": "One-line bullet point with scale/scope",
                "bp3": "One-line bullet point with compliance/security impact"
            }}
        }}

        Remember: BREVITY IS CRITICAL. Each bullet point should be 15-25 words maximum.
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