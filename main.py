import os
import sys
import json
import re
from pathlib import Path
from typing import Optional

from job_parser import AIJobParser
from experience_adapter import ExperienceAdapter
from cv_customizer import CVCustomizer
from cover_letter_generator import CoverLetterGenerator

def load_config():
    """Load configuration from config/settings.json"""
    config_dir = Path("config")
    config_file = config_dir / "settings.json"
    
    # Create config directory if it doesn't exist
    config_dir.mkdir(exist_ok=True)
    
    # Check if config file exists
    if not config_file.exists():
        # Update config with enhanced skills structure
        sample_config = {
            "api_key": "your_claude_api_key_here",
            "template_path": "templates/template.docx",
            "original_cv_path": "data/orig.docx",
            "context_path": "data/project_context.json",
            "input_file": "data/input.json",
            "skills_config": {
                "my_skills": {
                    "tier_1_core": [
                        "Python", "Java", "JavaScript", "SQL", "AWS"
                    ],
                    "tier_2_major": [
                        "Docker", "Terraform", "PostgreSQL", "Linux", "Git",
                        "Flask", "Django", "FastAPI", "Spring Boot", "jQuery"
                    ],
                    "tier_3_specialist": [
                        "RDS", "DynamoDB", "Lambda", "SQS", "S3", "EC2", "ECS",
                        "MySQL", "MongoDB", "Redis", "Elasticsearch", "SQLite",
                        "Jenkins", "GitHub Actions", "Prometheus", "Grafana",
                        "ELK Stack", "CloudWatch"
                    ],
                    "tier_4_tools": [
                        "Pandas", "NumPy", "Ubuntu", "CentOS", "macOS", "Windows Server",
                        "Bash", "PowerShell", "Vim", "Jira", "Confluence", "pytest",
                        "Cypress", "Jest", "Postman", "Cucumber", "Amazon SQS",
                        "Scikit-learn", "OWASP", "Wireshark", "Nmap", "OAuth",
                        "SAML", "Vault (HashiCorp)", "JWT", "ETL", "Data Migration",
                        "System Administration", "Performance Optimization",
                        "Monitoring", "Security", "Apache Airflow", "OpenSearch"
                    ]
                },
                "programming_languages": [
                    "Python", "Java", "JavaScript", "SQL"
                ],
                "blacklisted_skills": [
                    "TypeScript", "Go", "Rust", "C++", "C#", "C", "PHP", "Ruby",
                    "Scala", "Kotlin", "Swift", "R", "MATLAB", "Express.js", "React",
                    "Angular", "Node.js", "Bootstrap", "Apache Spark", "Apache Kafka",
                    "Jupyter", "Matplotlib", "Seaborn", "Plotly", "Apache Beam",
                    "Dask", "Polars", "Cassandra", "InfluxDB", "Neo4j", "Oracle",
                    "SQL Server", "Azure", "Google Cloud Platform", "Azure Functions",
                    "Google Cloud Functions", "Kubernetes", "Ansible", "GitLab CI/CD",
                    "CircleCI", "Helm", "Vagrant", "DataDog", "New Relic", "SVN",
                    "JUnit", "Selenium", "Mocha", "SonarQube", "TestNG", "Apache Kafka",
                    "RabbitMQ", "Apache ActiveMQ", "Azure Service Bus", "Google Pub/Sub",
                    "Apache Pulsar", "Splunk", "Jaeger", "Zipkin", "TensorFlow",
                    "PyTorch", "Keras", "OpenCV", "NLTK", "spaCy", "Transformers",
                    "MLflow", "Kubeflow", "Databricks", "Snowflake", "DBT",
                    "Great Expectations", "Apache NiFi", "Talend", "Pentaho",
                    "Machine Learning", "Deep Learning", "Metasploit", "Burp Suite",
                    "Nessus"
                ]
            }
        }
        
        with open(config_file, 'w') as f:
            json.dump(sample_config, f, indent=2)
        
        print(f"📁 Created config directory: {config_dir}")
        print(f"⚙️  Created sample config file: {config_file}")
        print("\n🔑 Please update config/settings.json with your Claude API key")
        print("   Get your API key from: https://console.anthropic.com/")
        return None
    
    # Load existing config
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        # Validate required fields
        if not config.get('api_key') or config['api_key'] == 'your_claude_api_key_here':
            print("❌ Please set your Claude API key in config/settings.json")
            print("   Get your API key from: https://console.anthropic.com/")
            return None
        
        return config
        
    except json.JSONDecodeError as e:
        print(f"❌ Error parsing config file: {e}")
        print("   Please check config/settings.json for valid JSON format")
        return None
    except Exception as e:
        print(f"❌ Error loading config: {e}")
        return None

class EnhancedCVCustomizer:
    def __init__(self, template_path: str, api_key: str, original_cv_path: str = None, context_path: str = None, skills_config: dict = None):
        """Enhanced CV customizer with AI integration"""
        self.cv_customizer = CVCustomizer(template_path)
        self.job_parser = AIJobParser(api_key)
        self.experience_adapter = ExperienceAdapter(api_key, original_cv_path, context_path, skills_config or {})
        self.cover_letter_generator = CoverLetterGenerator(api_key)
    
    def create_cv_from_input_file(self, input_file: str, output_name: Optional[str] = None):
        """Main method: provide input file with job description, get customized CV and cover letter"""
        
        # Read original job text for logging
        with open(input_file, 'r', encoding='utf-8') as f:
            original_job_text = f.read().strip()
        
        print("🤖 Parsing job description with AI...")
        job_info = self.job_parser.parse_from_file(input_file)
        
        print(f"📋 Parsed job: {job_info.job_title} at {job_info.company_name}")
        print(f"🎯 Required skills: {', '.join(job_info.required_skills[:5])}{'...' if len(job_info.required_skills) > 5 else ''}")
        
        print("🔧 Adapting your experience to match job requirements...")
        adapted_experience = self.experience_adapter.adapt_experience_to_job(job_info)
        
        # Generate output name if not provided
        if not output_name:
            company_safe = re.sub(r'[^\w\s-]', '', job_info.company_name or 'Company').replace(' ', '_')
            title_safe = re.sub(r'[^\w\s-]', '', job_info.job_title or 'Position').replace(' ', '_')
            output_name = f"CV_{company_safe}_{title_safe}"
        
        print("📝 Generating customized CV...")
        docx_path, pdf_path, execution_folder = self.cv_customizer.customize_cv(
            adapted_experience, output_name, job_info, original_job_text
        )
        
        # Smart optimization: Check if we can enhance the CV with more detail
        print("🔍 Analyzing page space usage...")
        page_usage = self.cv_customizer.estimate_page_usage(adapted_experience)
        print(f"📊 Estimated page usage: {page_usage:.1%}")
        
        if page_usage < 0.75:  # If using less than 75% of page
            print("🚀 Extra space available - regenerating with enhanced detail for maximum impact...")
            try:
                enhanced_experience = self.experience_adapter.regenerate_with_enhanced_detail(job_info, adapted_experience)
                
                # Regenerate CV with enhanced content
                enhanced_docx_path, enhanced_pdf_path, _ = self.cv_customizer.customize_cv(
                    enhanced_experience, output_name, job_info, original_job_text
                )
                
                # Verify enhanced version doesn't exceed page limit
                enhanced_page_usage = self.cv_customizer.estimate_page_usage(enhanced_experience)
                print(f"📈 Enhanced page usage: {enhanced_page_usage:.1%}")
                
                if enhanced_page_usage <= 1.0:  # If still within page limit
                    print("✨ Enhanced CV generated with maximum detail!")
                    docx_path, pdf_path = enhanced_docx_path, enhanced_pdf_path
                else:
                    print("⚠️  Enhanced version too long, using original optimized version")
                    
            except Exception as e:
                print(f"⚠️  Enhancement failed: {e}, using original version")
        else:
            print("✓ Optimal page usage achieved with current detail level")
        
        # Generate cover letter
        print("💌 Generating tailored cover letter...")
        try:
            cover_letter_filename = f"Drew_Gillies_Cover_Letter_{job_info.company_name.replace(' ', '_')}.pdf"
            cover_letter_path = f"{execution_folder}/{cover_letter_filename}"
            
            cover_letter_text = self.cover_letter_generator.generate_cover_letter(job_info, cover_letter_path)
            print(f"✓ Cover letter generated: {cover_letter_path}")
            
        except Exception as e:
            print(f"⚠️  Cover letter generation failed: {e}")
            print("   CV files are still available")
            cover_letter_path = None
        
        return docx_path, pdf_path, cover_letter_path
    
    def create_cv_from_job_description(self, job_description: str, output_name: Optional[str] = None):
        """Alternative method: provide job description text directly"""
        
        print("🤖 Parsing job description with AI...")
        job_info = self.job_parser.parse_job_description(job_description)
        
        print(f"📋 Parsed job: {job_info.job_title} at {job_info.company_name}")
        
        print("🔧 Adapting your experience to match job requirements...")
        adapted_experience = self.experience_adapter.adapt_experience_to_job(job_info)
        
        # Generate output name if not provided
        if not output_name:
            company_safe = re.sub(r'[^\w\s-]', '', job_info.company_name or 'Company').replace(' ', '_')
            title_safe = re.sub(r'[^\w\s-]', '', job_info.job_title or 'Position').replace(' ', '_')
            output_name = f"CV_{company_safe}_{title_safe}"
        
        print("📝 Generating customized CV...")
        return self.cv_customizer.customize_cv(adapted_experience, output_name)

def create_sample_input_file(input_file_path: str):
    """Create a sample input file for user reference"""
    sample_job_text = """Senior Python Developer - TechCorp London

About the Role:
We're looking for a Senior Python Developer to join our growing fintech team in London. You'll be working on our core trading platform that processes millions of transactions daily.

Requirements:
- 3+ years of experience with Python
- Strong experience with AWS (Lambda, RDS, DynamoDB, SQS)
- Experience with SQL and database optimization
- Docker and containerization experience
- Experience with microservices architecture
- Financial services experience preferred
- Knowledge of event-driven systems

Responsibilities:
- Develop and maintain Python applications for trading systems
- Optimize database performance and queries
- Build scalable microservices on AWS
- Collaborate with cross-functional teams
- Implement automated testing and CI/CD pipelines
- Ensure system reliability and performance

What we offer:
- Competitive salary (£80k-120k)
- Remote-friendly work environment
- Stock options
- Comprehensive health benefits

Location: London, UK (Hybrid - 3 days in office)"""
    
    with open(input_file_path, 'w') as f:
        f.write(sample_job_text)
    
    print(f"📄 Created sample {input_file_path}")
    print("   You can now paste any raw job description directly into this file")
    print("   No JSON formatting needed - just paste the entire job posting!")

def main():
    """Main application entry point"""
    
    print("🚀 AI-Powered CV Customizer")
    print("=" * 40)
    
    # Load configuration
    config = load_config()
    if not config:
        return
    
    # Get paths from config
    CLAUDE_API_KEY = config['api_key']
    TEMPLATE_PATH = config.get('template_path', 'templates/template.docx')
    ORIGINAL_CV_PATH = config.get('original_cv_path', 'data/orig.docx')
    CONTEXT_PATH = config.get('context_path', 'data/project_context.json')
    INPUT_FILE = config.get('input_file', 'data/input.json')
    
    # Check for required files
    missing_files = []
    if not Path(TEMPLATE_PATH).exists():
        missing_files.append(TEMPLATE_PATH)
    if not Path(ORIGINAL_CV_PATH).exists():
        missing_files.append(ORIGINAL_CV_PATH)
    
    if missing_files:
        print(f"❌ Missing required files: {', '.join(missing_files)}")
        print("   Make sure you have:")
        print(f"   - {TEMPLATE_PATH}: Your CV template with placeholders")
        print(f"   - {ORIGINAL_CV_PATH}: Your original CV for reference")
        sys.exit(1)
    
    # Check for input file
    if not Path(INPUT_FILE).exists():
        print(f"📝 Input file '{INPUT_FILE}' not found.")
        print("Creating a sample file for you...")
        # Create data directory if it doesn't exist
        Path(INPUT_FILE).parent.mkdir(exist_ok=True)
        create_sample_input_file(INPUT_FILE)
        print("\n💡 Next steps:")
        print(f"   1. Open '{INPUT_FILE}' in any text editor")
        print("   2. Replace the content with your job description (raw text)")
        print("   3. Run the script again: python main.py")
        return
    
    # Initialize the customizer
    print("🔧 Initializing AI CV customizer...")
    skills_config = config.get('skills_config', {})
    customizer = EnhancedCVCustomizer(
        TEMPLATE_PATH, 
        CLAUDE_API_KEY, 
        ORIGINAL_CV_PATH, 
        CONTEXT_PATH if Path(CONTEXT_PATH).exists() else None,
        skills_config
    )
    
    if Path(CONTEXT_PATH).exists():
        print("✓ Loaded additional project context")
    else:
        print("ℹ️  No project context file found (optional)")
    
    try:
        # Generate customized CV from input file
        docx_path, pdf_path, cover_letter_path = customizer.create_cv_from_input_file(INPUT_FILE)
        
        print("\n" + "=" * 40)
        print("✅ CV customization complete!")
        print(f"📄 Word document: {docx_path}")
        if pdf_path:
            print(f"📑 PDF document: {pdf_path}")
        if cover_letter_path:
            print(f"💌 Cover letter: {cover_letter_path}")
        
        print("\n💡 Tips:")
        print("   - Review the generated CV and cover letter before sending")
        print("   - All files are ready for direct upload to job sites")
        print("   - Original job description is saved for reference")
        print("   - Paste any raw job description into input file (no formatting needed)")
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("\nTroubleshooting:")
        print("1. Check that your input file has valid job description")
        print("2. Verify API key is set correctly in config/settings.json")
        print("3. Ensure template.docx has the correct placeholders")
        print("4. Make sure you have internet connection for AI processing")
        print("5. Install missing dependencies: pip install reportlab")

if __name__ == "__main__":
    main()