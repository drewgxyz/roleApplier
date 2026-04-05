import anthropic
import json
import re
from dataclasses import dataclass
from typing import List

# Model for job parsing - use Sonnet for accuracy on skill extraction
CLAUDE_MODEL = "claude-sonnet-4-6"

@dataclass
class JobInfo:
    """Structured job information extracted by AI"""
    job_title: str
    company_name: str
    location: str
    required_skills: List[str]
    preferred_skills: List[str]
    years_experience: str
    key_responsibilities: List[str]
    industry: str
    remote_policy: str

class AIJobParser:
    def __init__(self, api_key: str):
        """Initialize with Claude API key"""
        self.client = anthropic.Anthropic(api_key=api_key)
    
    def parse_job_description(self, job_text: str) -> JobInfo:
        """Parse raw job description text using Claude"""
        
        prompt = f"""
        Please analyze this job description and extract key information in JSON format.
        
        This could be a raw dump from a job website, so please extract the relevant job information and ignore any website navigation, ads, or unrelated content.

        Job Description Text:
        {job_text}

        Extract and return ONLY a JSON object with these fields:
        - job_title: The job title/position
        - company_name: Company name (if mentioned, otherwise "Unknown Company")
        - location: Location/city (extract from text, use "Remote" if remote work mentioned, "Not specified" if unclear)
        - required_skills: Array of technical skills, programming languages, frameworks, tools mentioned as required/essential
        - preferred_skills: Array of nice-to-have skills mentioned (or empty array if none specified)
        - years_experience: Experience level required (e.g. "3-5 years", "Senior level", "Junior", "Mid-level")
        - key_responsibilities: Array of main job responsibilities/duties (3-5 key ones)
        - industry: Industry sector (e.g. "Fintech", "Healthcare", "E-commerce", "Technology", "Consulting")
        - remote_policy: "Remote", "Hybrid", "On-site", or "Not specified"

        Focus on extracting specific technologies, frameworks, and tools mentioned.
        Ignore any website navigation, footer content, or unrelated text.
        """

        try:
            response = self.client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=1500,
                messages=[{"role": "user", "content": prompt}]
            )
            
            # Extract JSON from response
            response_text = response.content[0].text
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            
            if json_match:
                job_data = json.loads(json_match.group())
                
                # Filter to only expected fields to avoid unexpected keyword arguments
                expected_fields = {
                    'job_title', 'company_name', 'location', 'required_skills', 
                    'preferred_skills', 'years_experience', 'key_responsibilities', 
                    'industry', 'remote_policy'
                }
                
                # Create filtered dict with only expected fields
                filtered_data = {}
                for field in expected_fields:
                    if field in job_data:
                        filtered_data[field] = job_data[field]
                    else:
                        # Provide defaults for missing fields
                        if field == 'required_skills':
                            filtered_data[field] = []
                        elif field == 'preferred_skills':
                            filtered_data[field] = []
                        elif field == 'key_responsibilities':
                            filtered_data[field] = []
                        elif field == 'company_name':
                            filtered_data[field] = "Unknown Company"
                        elif field == 'location':
                            filtered_data[field] = "Not specified"
                        elif field == 'years_experience':
                            filtered_data[field] = "Not specified"
                        elif field == 'industry':
                            filtered_data[field] = "Technology"
                        elif field == 'remote_policy':
                            filtered_data[field] = "Not specified"
                        elif field == 'job_title':
                            filtered_data[field] = "Software Role"
                        else:
                            filtered_data[field] = ""
                
                return JobInfo(**filtered_data)
            else:
                raise ValueError("Could not extract JSON from AI response")
                
        except Exception as e:
            print(f"Error parsing job description: {e}")
            raise

    def parse_from_file(self, input_file: str) -> JobInfo:
        """Parse job description from input file (supports both JSON and raw text)"""
        try:
            with open(input_file, 'r', encoding='utf-8') as f:
                content = f.read().strip()
            
            if not content:
                raise ValueError("Input file is empty")
            
            print(f"📄 Read {len(content)} characters from {input_file}")
            print(f"🔍 First 100 chars: {content[:100]}...")
            
            # Since your file doesn't start with {, it should be treated as raw text
            job_text = content
            
            # Only try JSON parsing if it clearly looks like JSON
            if content.startswith('{') and content.endswith('}'):
                try:
                    print("🔧 Attempting JSON parse...")
                    data = json.loads(content)
                    job_text = data.get('job_description', content)
                    print("✓ Successfully parsed as JSON")
                except json.JSONDecodeError as e:
                    print(f"⚠️  JSON parse failed: {e}, treating as raw text")
                    job_text = content
            else:
                print("📝 Treating as raw text (not JSON)")
            
            return self.parse_job_description(job_text)
            
        except FileNotFoundError:
            raise ValueError(f"Input file '{input_file}' not found")
        except Exception as e:
            print(f"Error reading job description file: {e}")
            print(f"Error type: {type(e).__name__}")
            raise