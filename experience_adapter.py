import anthropic
import json
import re
from pathlib import Path
from typing import Dict
from docx import Document
from job_parser import JobInfo

# Model for experience adaptation - use Sonnet for quality writing
CLAUDE_MODEL = "claude-sonnet-4-6"

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


    def _get_adaptive_content_strategy(self, current_strategy: dict) -> dict:
        """Create a slightly less detailed content strategy to reduce page usage."""
        new_strategy = current_strategy.copy()
        print("🔧 Adapting content strategy to reduce length...")

        # Reduce content in a prioritized order
        if new_strategy['expertise_count'] > 12:
            new_strategy['expertise_count'] -= 2
            print(f"   - Reduced expertise count to {new_strategy['expertise_count']}")
        elif new_strategy['trp_detail'] == 'MAXIMUM':
            new_strategy['trp_detail'] = 'HIGH'
            print("   - Reduced T. Rowe Price detail to HIGH")
        elif new_strategy['aws_detail'] == 'HIGH':
            new_strategy['aws_detail'] = 'MEDIUM'
            print("   - Reduced AWS detail to MEDIUM")
        elif new_strategy['bio_sentences'] > 2:
            new_strategy['bio_sentences'] -= 1
            print(f"   - Reduced bio sentences to {new_strategy['bio_sentences']}")
        else:
            # Fallback if everything is already minimal
            new_strategy['expertise_count'] -= 1
            print(f"   - Fallback: Reduced expertise count to {new_strategy['expertise_count']}")
            
        return new_strategy
    
    def _extract_cv_data(self, cv_path: str) -> Dict:
        """Extract experience data from original CV document"""
        if not cv_path or not Path(cv_path).exists():
            return {}
            
        try:
            doc = Document(cv_path)
            cv_text = "\n".join([para.text for para in doc.paragraphs])
            
            # Extract the original experience data structure
            return {
                "compare_the_market": {
                    "original_skills": "Python, LangChain, LangGraph, Redis, PostgreSQL, AWS ECS, AWS EFS",
                    "projects": {
                        "bp1": "Led architecture of AI-powered product requirements pipeline, automatically converting PRDs into JIRA-ready engineering tickets using Python, LangGraph, LangChain, Redis, PostgreSQL, AWS ECS & EFS, reducing manual workload by 80%+",
                        "bp2": "Developed internal AI-driven code review and SDLC automation platform used across 300+ engineers, increasing PR throughput by 75% and accelerating iteration cycles",
                        "bp3": "Engineered distributed AI systems integrating caching (Redis), persistent state (PostgreSQL), and event-driven processing patterns, improving performance and cost efficiency of AI workloads",
                        "bp4": "Led AI adoption across engineering teams through workshops, hackathons, and internal talks, enabling engineers to integrate AI into production workflows"
                    }
                },
                "t_rowe_price": {
                    "original_skills": "Python, Java, AWS, RDS, Terraform",
                    "projects": {
                        "bp1": "Led architecture and development of production-grade Python data migration tool for syncing complex relational data across DEV/STAGE/PROD environments with rollback safety, referential integrity validation, and automated testing",
                        "bp2": "Redesigned legacy application with database performance issues, implementing scalable event-driven architecture on AWS using Lambda, SQS, Python, and Java, achieving 60% performance improvement",
                        "bp3": "Architected disaster recovery strategies for legacy core infrastructure, rebuilding critical Java services in Python to enable seamless DR failover and establishing automated backup systems across US-East-2 region",
                        "bp4": "Developed high-performance data loaders for internal directory system, integrating Active Directory data into RDS and OpenSearch to optimize search functionality"
                    }
                },
                "aws": {
                    "original_skills": "Python, AWS, Linux System Administration",
                    "projects": {
                        "bp1": "Optimized Region Build deployment process, reducing required time by 40-55% across 15+ Service Catalog services and pipelines",
                        "bp2": "Deployed Service Catalog services across 5 newly launched AWS Regions (UAE, Melbourne, Spain, Zurich, Hyderabad) supporting global expansion",
                        "bp3": "Led security escalation response managing 2,400+ hosts, implementing automated security patching pipelines"
                    }
                },
                "bio": "Software Engineer with 3.5 years experience building AI-native automation systems, enterprise data pipelines, and cloud infrastructure. Currently leading AI product automation at Compare the Market, previously delivered production-grade migration tools at T. Rowe Price and global infrastructure deployments at AWS. Key expertise in Python, LangChain, LangGraph, AWS, and distributed systems."
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

        # Sanitize the job title to remove forbidden keywords before sending to the AI
        forbidden_keywords = r'\b(Senior|Lead|Principal|Staff|I|II|III|IV|DevOps)\b'
        sanitized_title = re.sub(forbidden_keywords, '', job_info.job_title, flags=re.IGNORECASE).strip()
        
        prompt = f"""
        I need to customize my CV for this specific job opportunity. I want to keep the exact same structure but adapt the content to highlight relevant skills and experiences.

        TARGET JOB:
        - Position: {sanitized_title} at {job_info.company_name}
        - Required Skills: {', '.join(job_info.required_skills)}
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
        2. Bio: Use exactly {content_strategy['bio_sentences']} sentences. NEVER CALL MY OWN ROLE IN THE BIO SOMETHING I AM NOT, i.e Senior Software Developer. I can be: Software Engineer, Software Developer, Backend Engineer
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
        - HIGH: 30-40 words, include specific technologies, metrics, and technical details
        - MEDIUM: 20-30 words, include key technologies and quantified outcomes  
        - LOW: 15-25 words, focus on impact and relevant technologies
        - MINIMAL: 10-15 words, essential impact only

        BULLET POINT REQUIREMENTS:
        - ALWAYS mention specific technologies from the tech stacks in bullet points
        - Include deliverable business value and quantified results when space allows
        - Use action verbs and technical specificity
        - Prioritize content relevance to the target job
        - If detail level is HIGH, include multiple technologies per bullet point
        - If detail level is LOW/MINIMAL, focus on most relevant single technology

        Return ONLY a JSON object with these exact fields:
        {{
            "bio": "Updated bio paragraph with exactly {content_strategy['bio_sentences']} sentences - NEVER use 'Senior' or inflated job titles",
            "expertise": ["Prioritized list: programming languages first, then job-relevant skills in order of importance - exactly {content_strategy['expertise_count']} skills total"],
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

        EXPERTISE SECTION PRIORITIZATION (CRITICAL):
        1. ALWAYS start with programming languages: Python, Java, JavaScript (in that order if relevant to job)
        2. Then add job-relevant frameworks/technologies in order of importance to this specific role
        3. Then add cloud platforms (AWS, etc.) if mentioned in job requirements
        4. Then add databases and tools in order of job relevance
        5. Fill remaining slots with most relevant skills from lower tiers
        
        Examples:
        - Python/Flask job: ["Python", "Java", "Flask", "FastAPI", "AWS", "PostgreSQL", "Docker", "Redis", ...]
        - Java/Spring job: ["Java", "Python", "Spring Boot", "AWS", "PostgreSQL", "Docker", ...]
        - Data Engineering job: ["Python", "SQL", "AWS", "PostgreSQL", "Docker", "ETL", "Apache Airflow", ...]
        - DevOps job: ["Python", "AWS", "Docker", "Terraform", "Jenkins", "Linux", ...]

        CRITICAL: Follow the exact content strategy provided. Technologies in tech stacks MUST appear in bullet points.
        NEVER inflate job titles - use only "Software Engineer", "Software Developer", or "Mid-level" versions.
        Adjust content density based on detail levels to ensure 1-page fit.
        """

        try:
            response = self.client.messages.create(
                model=CLAUDE_MODEL,
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
    

    def regenerate_with_enhanced_detail(self, job_info: JobInfo, current_data: Dict, strategy: dict = None) -> Dict:
        """
        Regenerate CV with enhanced detail levels for maximum impact.
        Accepts an optional strategy to be used by the adaptive loop.
        """
        
        # Use the provided strategy if it exists; otherwise, generate the default enhanced one.
        # This makes the method compatible with the new adaptive loop in main.py.
        if strategy:
            enhanced_strategy = strategy
            # We still need relevant skills and score for the prompt context.
            relevant_skills = self._filter_job_skills_with_priority(job_info.required_skills + job_info.preferred_skills)
            relevance_score = self._assess_job_relevance(job_info, relevant_skills)
        else:
            # Fallback for old behavior if no strategy is passed
            relevant_skills = self._filter_job_skills_with_priority(job_info.required_skills + job_info.preferred_skills)
            relevance_score = self._assess_job_relevance(job_info, relevant_skills)
            enhanced_strategy = self._get_enhanced_content_strategy(relevance_score)
        
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
        
        CONTENT STRATEGY:
        Job Relevance Score: {relevance_score}/10
        Strategy Name: {enhanced_strategy['name']}
        Bio Length: {enhanced_strategy['bio_sentences']} sentences
        T. Rowe Price Detail Level: {enhanced_strategy['trp_detail']}
        AWS Detail Level: {enhanced_strategy['aws_detail']}
        Expertise Skills Count: {enhanced_strategy['expertise_count']}
        """
        
        # Include additional context in the prompt
        context_section = ""
        if self.additional_context:
            context_section = f"""
        ADDITIONAL PROJECT CONTEXT (use extensively for detail):
        {json.dumps(self.additional_context, indent=2)}
        """
        
        # Sanitize the job title to remove forbidden keywords before sending to the AI
        forbidden_keywords = r'\b(Senior|Lead|Principal|Staff|I|II|III|IV|DevOps)\b'
        sanitized_title = re.sub(forbidden_keywords, '', job_info.job_title, flags=re.IGNORECASE).strip()

        # The prompt remains the same, but now it uses the passed-in strategy
        prompt = f"""
        ENHANCED CV GENERATION - ADAPTIVE STRATEGY

        I need to create the STRONGEST possible CV for this specific job, adhering to the provided adaptive content strategy.

        TARGET JOB:
        - Position: {sanitized_title} at {job_info.company_name}
        - Required Skills: {', '.join(job_info.required_skills)}

        {skills_section}

        MY CURRENT EXPERIENCE (to adapt from):
        {json.dumps(self.original_cv_data, indent=2)}
        
        {context_section}

        CRITICAL INSTRUCTIONS - ADAPT TO THE PROVIDED STRATEGY:
        1. Bio: Use exactly {enhanced_strategy['bio_sentences']} sentences. NEVER CALL MY OWN ROLE IN THE BIO SOMETHING I AM NOT, i.e Senior Software Developer. I can be: Software Engineer, Software Developer, Backend Engineer
        2. T. Rowe Price bullets: Use '{enhanced_strategy['trp_detail']}' detail level.
        3. AWS bullets: Use '{enhanced_strategy['aws_detail']}' detail level.
        4. Expertise: Include exactly {enhanced_strategy['expertise_count']} skills.
        5. CRITICAL: For the "expertise" list, YOU MUST ONLY USE skills from the 'MY ACTUAL SKILLS' or 'JOB RELEVANT SKILLS' lists provided earlier. DO NOT invent conceptual categories like 'Python Development'. Use the actual skill names like 'Python', 'FastAPI', etc.

        Return ONLY a JSON object with this exact nested structure. This is non-negotiable.
        {{
            "bio": "The updated bio paragraph.",
            "expertise": ["ULTRA-CRITICAL: A list of exactly {enhanced_strategy['expertise_count']} skills. YOU MUST ONLY use skills from the 'MY ACTUAL SKILLS' list (e.g., 'Python', 'FastAPI', 'Docker'). DO NOT invent conceptual categories like 'Python Development' or 'Database Architecture' under any circumstances."],
            "t": {{
                "skills": "ULTRA-CRITICAL: A single comma-separated STRING of technologies (e.g., 'Python, AWS, Docker'). The value for this key MUST be a string, NOT a list.",
                "bp1": "The first bullet point for T. Rowe Price.",
                "bp2": "The second bullet point for T. Rowe Price.",
                "bp3": "The third bullet point for T. Rowe Price.",
                "bp4": "The fourth bullet point for T. Rowe Price."
            }},
            "a": {{
                "skills": "ULTRA-CRITICAL: A single comma-separated STRING of technologies (e.g., 'Python, AWS, Docker'). The value for this key MUST be a string, NOT a list.",
                "bp1": "The first bullet point for AWS.",
                "bp2": "The second bullet point for AWS.",
                "bp3": "The third bullet point for AWS."
            }}
        }}
        """
        
        try:
            response = self.client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=2500,
                messages=[{"role": "user", "content": prompt}]
            )
            
            response_text = response.content[0].text
            # Clean up potential trailing commas before parsing
            response_text = re.sub(r',\s*([\}\]])', r'\1', response_text)
            
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            
            if json_match:
                return json.loads(json_match.group())
            else:
                raise ValueError("Could not extract JSON from AI response")
                
        except Exception as e:
            print(f"Error in enhanced regeneration: {e}")
            # Return original data as fallback if something goes wrong
            return current_data
    
    def _validate_bullet_point_length(self, cv_data: Dict) -> Dict:
        """Validate T. Rowe Price bullet points meet minimum length requirements"""
        issues = []
        MIN_WORD_COUNT = 50

        if 't' in cv_data and isinstance(cv_data['t'], dict):
            for bp_key in ['bp1', 'bp2', 'bp3', 'bp4']:
                if bp_key in cv_data['t']:
                    bp_text = cv_data['t'][bp_key]
                    word_count = len(bp_text.split())
                    
                    if word_count < MIN_WORD_COUNT:
                        issues.append(f"T. Rowe Price {bp_key}: {word_count} words (minimum {MIN_WORD_COUNT} required)")
        
        return {
            'has_issues': len(issues) > 0,
            'issues': issues,
            'total_issues': len(issues)
        }
    
    def _regenerate_with_length_enforcement(self, job_info: JobInfo, current_data: Dict, retry_count: int = 0) -> Dict:
        """Regenerate with strict length enforcement for T. Rowe Price bullet points"""
        
        if retry_count >= 2:  # Max 2 retries
            print("⚠️  Maximum retries reached, using current version")
            return current_data
        
        print(f"🔄 Retry {retry_count + 1}: Regenerating with STRICT length requirements...")
        
        # Get enhanced strategy
        relevant_skills = self._filter_job_skills_with_priority(job_info.required_skills + job_info.preferred_skills)
        relevance_score = self._assess_job_relevance(job_info, relevant_skills)
        enhanced_strategy = self._get_enhanced_content_strategy(relevance_score)
        
        # Ultra-strict prompt for length enforcement
        strict_prompt = f"""
        CRITICAL FAILURE ANALYSIS: The previous attempt failed because the bullet points for T. Rowe Price were too short. Your task is to fix this. This is a non-negotiable instruction. 

        TARGET JOB:
        - Position: {job_info.job_title}
        - Required Skills: {', '.join(job_info.required_skills)}

        DETAILED PROJECT CONTEXT (USE THIS EXTENSIVELY):
        {json.dumps(self.additional_context, indent=2) if self.additional_context else ""}

        MANDATORY INSTRUCTIONS FOR T. ROWE PRICE BULLET POINTS:
        1.  **WORD COUNT: 65-75 words. EACH. This is an absolute, non-negotiable requirement.** Do not generate anything shorter.
        2.  **TECHNICAL SYNTHESIS:** To achieve this length, you MUST synthesize information. For each bullet point, combine details from the 'detailed_description', 'technical_stack', 'challenges_solved', and 'business_impact' sections of the provided context.
        3.  **TECHNOLOGY INTEGRATION:** You MUST mention 4-5 specific technologies from the tech stack in each bullet point. Weave them into the narrative naturally.
        4.  **EXAMPLE OF SYNTHESIS:** For 'bp1', you could start with the description, then mention the 'SQLAlchemy ORM', 'PostgreSQL', and 'Docker' from the stack, explain how they solved the 'complex foreign key relationships' challenge, and quantify the '95% reduction in manual migration time' as the impact.

        ---
        PERFECT EXAMPLE BULLET POINT (75 words):
        "Architected and led the development of a production-grade Python data migration tool using the FastAPI framework, leveraging SQLAlchemy for complex relational data mapping in PostgreSQL and utilizing Redis for caching to ensure rollback safety. This tool, containerized with Docker and deployed on AWS, automated the synchronization of data across DEV/STAGE/PROD environments, reducing manual migration time by 95% and eliminating data integrity errors through comprehensive automated validation scripts, ensuring referential integrity for critical financial reporting systems."
        ---

        Now, regenerate the ENTIRE JSON object. Adhere strictly to the 65-75 word count for every T. Rowe Price bullet point.

        Return ONLY a JSON object with this exact nested structure. This is non-negotiable.
        {{
            "bio": "The updated bio paragraph with exactly {enhanced_strategy['bio_sentences']} sentences.",
            "expertise": ["CRITICAL: A list of exactly {enhanced_strategy['expertise_count']} skills. YOU MUST ONLY USE skills from the 'MY ACTUAL SKILLS' list. DO NOT invent conceptual categories like 'Python Development'. Use the actual skill names like 'Python', 'FastAPI', 'Docker', etc."],
            "t": {{
                "skills": "CRITICAL: A single comma-separated string of technologies (e.g., 'Python, AWS, Docker'). DO NOT use a JSON list of strings.",
                "bp1": "The first bullet point for T. Rowe Price, written at '{enhanced_strategy['trp_detail']}' detail.",
                "bp2": "The second bullet point for T. Rowe Price, written at '{enhanced_strategy['trp_detail']}' detail.",
                "bp3": "The third bullet point for T. Rowe Price, written at '{enhanced_strategy['trp_detail']}' detail.",
                "bp4": "The fourth bullet point for T. Rowe Price, written at '{enhanced_strategy['trp_detail']}' detail."
            }},
            "a": {{
                "skills": "CRITICAL: A single comma-separated string of technologies (e.g., 'Python, AWS, Docker'). DO NOT use a JSON list of strings.",
                "bp1": "The first bullet point for AWS, written at '{enhanced_strategy['aws_detail']}' detail.",
                "bp2": "The second bullet point for AWS, written at '{enhanced_strategy['aws_detail']}' detail.",
                "bp3": "The third bullet point for AWS, written at '{enhanced_strategy['aws_detail']}' detail."
            }}
        }}
        """

        try:
            response = self.client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=1500,
                temperature=0.7,  # <--- ADD THIS LINE. Increases creativity and verbosity.
                messages=[{"role": "user", "content": strict_prompt}]
            )
            
            response_text = response.content[0].text
            
            # Clean up potential trailing commas before parsing
            response_text = re.sub(r',\s*([\}\]])', r'\1', response_text) # <--- ADD THIS LINE
            
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            
            if json_match:
                new_data = json.loads(json_match.group())
                
                # Validate the new data
                validation = self._validate_bullet_point_length(new_data)
                
                if validation['has_issues']:
                    print(f"❌ Length validation failed: {validation['total_issues']} issues")
                    for issue in validation['issues']:
                        print(f"   - {issue}")
                    
                    # Retry with stricter requirements
                    return self._regenerate_with_length_enforcement(job_info, current_data, retry_count + 1)
                else:
                    print("✅ All T. Rowe Price bullet points meet 70+ word requirement")
                    return new_data
            else:
                print("❌ Could not extract JSON from response")
                return self._regenerate_with_length_enforcement(job_info, current_data, retry_count + 1)
                
        except Exception as e:
            print(f"❌ Error in length enforcement: {e}")
            return self._regenerate_with_length_enforcement(job_info, current_data, retry_count + 1)
    
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
                'expertise_count': 14  # Increased from 12
            }
        elif relevance_score >= 6:
            # Medium-high relevance - balanced approach
            return {
                'name': 'MEDIUM_HIGH_RELEVANCE', 
                'bio_sentences': 3,
                'trp_detail': 'MEDIUM',
                'aws_detail': 'MEDIUM',
                'expertise_count': 12  # Increased from 10
            }
        elif relevance_score >= 4:
            # Medium relevance - conservative approach
            return {
                'name': 'MEDIUM_RELEVANCE',
                'bio_sentences': 3,
                'trp_detail': 'MEDIUM',
                'aws_detail': 'LOW',
                'expertise_count': 10  # Same as before
            }
        else:
            # Low relevance - minimal approach
            return {
                'name': 'LOW_RELEVANCE',
                'bio_sentences': 3,
                'trp_detail': 'LOW',
                'aws_detail': 'MINIMAL',
                'expertise_count': 8   # Same as before
            }
    
    def _get_enhanced_content_strategy(self, relevance_score: int) -> dict:
        """Get enhanced content strategy for maximum impact when space allows"""
        
        if relevance_score >= 8:
            # High relevance - maximum detail for best match
            return {
                'name': 'MAXIMUM_IMPACT_HIGH_RELEVANCE',
                'bio_sentences': 3,  # Shorter bio for more bullet space
                'trp_detail': 'MAXIMUM',
                'aws_detail': 'HIGH',
                'expertise_count': 14  # Increased from 14
            }
        elif relevance_score >= 6:
            # Medium-high relevance - enhanced balanced approach
            return {
                'name': 'ENHANCED_MEDIUM_HIGH_RELEVANCE', 
                'bio_sentences': 4,
                'trp_detail': 'HIGH',
                'aws_detail': 'HIGH',
                'expertise_count': 14  # Increased from 12
            }
        elif relevance_score >= 4:
            # Medium relevance - enhanced conservative approach
            return {
                'name': 'ENHANCED_MEDIUM_RELEVANCE',
                'bio_sentences': 4,
                'trp_detail': 'HIGH',
                'aws_detail': 'MEDIUM',
                'expertise_count': 13  # Increased from 11
            }
        else:
            # Low relevance - still enhanced from original
            return {
                'name': 'ENHANCED_LOW_RELEVANCE',
                'bio_sentences': 4,
                'trp_detail': 'MEDIUM',
                'aws_detail': 'LOW',
                'expertise_count': 12  # Increased from 9
            }