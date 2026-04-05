"""
SQLite Database for CV Generator Configuration
Stores skills, experiences, and settings
"""

import sqlite3
import json
from pathlib import Path
from datetime import datetime

DATABASE_PATH = 'cv_generator.db'


def get_db():
    """Get database connection"""
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Initialize database with tables"""
    conn = get_db()
    cursor = conn.cursor()
    
    # Skills table - approved skills with tiers
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS skills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            tier TEXT NOT NULL DEFAULT 'tier_4_tools',
            category TEXT,
            is_blacklisted BOOLEAN DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Skill aliases - for matching variations (e.g., "AWS" matches "Amazon Web Services")
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS skill_aliases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            skill_id INTEGER NOT NULL,
            alias TEXT NOT NULL,
            FOREIGN KEY (skill_id) REFERENCES skills(id) ON DELETE CASCADE
        )
    ''')
    
    # Experiences/Projects table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS experiences (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company TEXT NOT NULL,
            role TEXT NOT NULL,
            start_date TEXT,
            end_date TEXT,
            is_current BOOLEAN DEFAULT 0,
            order_index INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Bullet points for experiences - with mutable tech stacks
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bullet_points (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            experience_id INTEGER NOT NULL,
            base_description TEXT NOT NULL,
            tech_placeholder TEXT,
            order_index INTEGER DEFAULT 0,
            FOREIGN KEY (experience_id) REFERENCES experiences(id) ON DELETE CASCADE
        )
    ''')
    
    # Tech stack options for bullet points - which skills can be swapped in
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bullet_tech_options (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bullet_id INTEGER NOT NULL,
            skill_id INTEGER NOT NULL,
            is_primary BOOLEAN DEFAULT 0,
            FOREIGN KEY (bullet_id) REFERENCES bullet_points(id) ON DELETE CASCADE,
            FOREIGN KEY (skill_id) REFERENCES skills(id) ON DELETE CASCADE
        )
    ''')
    
    # Settings table for general config
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Bio templates
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bio_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            template TEXT NOT NULL,
            is_default BOOLEAN DEFAULT 0
        )
    ''')
    
    conn.commit()
    conn.close()


def seed_default_data():
    """Seed database with default skills and experiences"""
    conn = get_db()
    cursor = conn.cursor()
    
    # Check if already seeded
    cursor.execute('SELECT COUNT(*) FROM skills')
    if cursor.fetchone()[0] > 0:
        conn.close()
        return
    
    # Default skills by tier
    skills_data = {
        'tier_1_core': [
            ('Python', 'language'),
            ('Java', 'language'),
            ('JavaScript', 'language'),
            ('SQL', 'language'),
            ('AWS', 'cloud'),
        ],
        'tier_2_major': [
            ('Docker', 'devops'),
            ('Terraform', 'devops'),
            ('PostgreSQL', 'database'),
            ('Linux', 'os'),
            ('Git', 'tool'),
            ('Flask', 'framework'),
            ('Django', 'framework'),
            ('FastAPI', 'framework'),
            ('Spring Boot', 'framework'),
            ('jQuery', 'framework'),
        ],
        'tier_3_specialist': [
            ('RDS', 'aws'),
            ('DynamoDB', 'aws'),
            ('Lambda', 'aws'),
            ('SQS', 'aws'),
            ('S3', 'aws'),
            ('EC2', 'aws'),
            ('ECS', 'aws'),
            ('MySQL', 'database'),
            ('MongoDB', 'database'),
            ('Redis', 'database'),
            ('Elasticsearch', 'database'),
            ('SQLite', 'database'),
            ('Jenkins', 'devops'),
            ('GitHub Actions', 'devops'),
            ('CloudWatch', 'aws'),
        ],
        'tier_4_tools': [
            ('Pandas', 'library'),
            ('NumPy', 'library'),
            ('Ubuntu', 'os'),
            ('CentOS', 'os'),
            ('macOS', 'os'),
            ('Windows Server', 'os'),
            ('Bash', 'scripting'),
            ('PowerShell', 'scripting'),
            ('Vim', 'tool'),
            ('Jira', 'tool'),
            ('Confluence', 'tool'),
            ('pytest', 'testing'),
            ('Cypress', 'testing'),
            ('Jest', 'testing'),
            ('Postman', 'tool'),
            ('Cucumber', 'testing'),
            ('Amazon SQS', 'aws'),
            ('Scikit-learn', 'library'),
            ('OWASP', 'security'),
            ('Wireshark', 'security'),
            ('Nmap', 'security'),
            ('OAuth', 'security'),
            ('SAML', 'security'),
            ('Vault (HashiCorp)', 'security'),
            ('JWT', 'security'),
            ('ETL', 'concept'),
            ('Data Migration', 'concept'),
            ('System Administration', 'concept'),
            ('Performance Optimization', 'concept'),
            ('Monitoring', 'concept'),
            ('Security', 'concept'),
            ('Apache Airflow', 'tool'),
            ('OpenSearch', 'database'),
            ('Kafka', 'messaging'),
            ('RabbitMQ', 'messaging'),
        ],
    }
    
    # Insert skills
    for tier, skills in skills_data.items():
        for skill_name, category in skills:
            cursor.execute(
                'INSERT OR IGNORE INTO skills (name, tier, category) VALUES (?, ?, ?)',
                (skill_name, tier, category)
            )
    
    # Blacklisted skills
    blacklisted = [
        'TypeScript', 'Go', 'Rust', 'C++', 'C#', 'C', 'PHP', 'Ruby', 'Scala', 'Kotlin',
        'Swift', 'R', 'MATLAB', 'Express.js', 'React', 'Angular', 'Node.js', 'Bootstrap',
        'Apache Spark', 'Jupyter', 'Matplotlib', 'Seaborn', 'Plotly', 'Apache Beam',
        'Dask', 'Polars', 'Cassandra', 'InfluxDB', 'Neo4j', 'Oracle', 'SQL Server',
        'Azure', 'Google Cloud Platform', 'Azure Functions', 'Google Cloud Functions',
        'Kubernetes', 'Ansible', 'GitLab CI/CD', 'CircleCI', 'Helm', 'Vagrant',
        'DataDog', 'New Relic', 'SVN', 'JUnit', 'Selenium', 'Mocha', 'SonarQube',
        'TestNG', 'Azure Service Bus', 'Google Pub/Sub', 'Apache Pulsar', 'Splunk',
        'Jaeger', 'Zipkin', 'TensorFlow', 'PyTorch', 'Keras', 'OpenCV', 'NLTK',
        'spaCy', 'Transformers', 'MLflow', 'Kubeflow', 'Databricks', 'Snowflake',
        'DBT', 'Great Expectations', 'Apache NiFi', 'Talend', 'Pentaho',
        'Machine Learning', 'Deep Learning', 'Metasploit', 'Burp Suite', 'Nessus',
    ]
    
    for skill in blacklisted:
        cursor.execute(
            'INSERT OR IGNORE INTO skills (name, tier, category, is_blacklisted) VALUES (?, ?, ?, ?)',
            (skill, 'tier_4_tools', 'blacklisted', 1)
        )
    
    # Default experiences
    cursor.execute('''
        INSERT INTO experiences (company, role, start_date, end_date, is_current, order_index)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', ('T. Rowe Price', 'Software Engineer', '2023-06', None, 1, 1))
    trp_id = cursor.lastrowid
    
    cursor.execute('''
        INSERT INTO experiences (company, role, start_date, end_date, is_current, order_index)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', ('Amazon Web Services', 'Cloud Support Engineer', '2022-07', '2023-06', 0, 2))
    aws_id = cursor.lastrowid
    
    # T. Rowe Price bullet points with mutable tech
    trp_bullets = [
        ('Led architecture and development of production-grade {tech} data migration tool for syncing complex relational data across DEV/STAGE/PROD environments with rollback safety, referential integrity validation, and automated testing', 'tech'),
        ('Redesigned legacy application with database performance issues, implementing scalable event-driven architecture on {cloud} using {services}, achieving 60% performance improvement', 'cloud,services'),
        ('Architected disaster recovery strategies for legacy core infrastructure, rebuilding critical {lang1} services in {lang2} to enable seamless DR failover and establishing automated backup systems', 'lang1,lang2'),
        ('Developed high-performance data loaders for internal directory system, integrating Active Directory data into {database} and {search} to optimize search functionality and reduce response times', 'database,search'),
    ]
    
    for idx, (desc, placeholders) in enumerate(trp_bullets):
        cursor.execute(
            'INSERT INTO bullet_points (experience_id, base_description, tech_placeholder, order_index) VALUES (?, ?, ?, ?)',
            (trp_id, desc, placeholders, idx + 1)
        )
    
    # AWS bullet points
    aws_bullets = [
        ('Optimized Region Build deployment process, reducing required time by 40-55% across 15+ Service Catalog services and pipelines', None),
        ('Deployed Service Catalog services across 5 newly launched AWS Regions (UAE, Melbourne, Spain, Zurich, Hyderabad) supporting global expansion', None),
        ('Led security escalation response managing 2,400+ hosts, implementing automated security patching pipelines', None),
    ]
    
    for idx, (desc, placeholders) in enumerate(aws_bullets):
        cursor.execute(
            'INSERT INTO bullet_points (experience_id, base_description, tech_placeholder, order_index) VALUES (?, ?, ?, ?)',
            (aws_id, desc, placeholders, idx + 1)
        )
    
    # Default bio template
    cursor.execute('''
        INSERT INTO bio_templates (name, template, is_default)
        VALUES (?, ?, ?)
    ''', ('Default', 'Software Engineer with {years} years of experience delivering production-grade solutions. Proven track record in {domains}. Key expertise includes {skills}.', 1))
    
    # Default settings
    settings = {
        'user_name': 'Drew Gillies',
        'user_email': 'drew.gillies@hotmail.co.uk',
        'user_phone': '07950 298726',
        'user_location': 'London, UK',
        'user_linkedin': 'linkedin.com/in/drew-gillies',
        'years_experience': '2.5',
        'education': 'BSc Cyber Security from Warwick University (2022)',
        'expertise_count': '14',
        'bio_sentences': '3',
    }
    
    for key, value in settings.items():
        cursor.execute(
            'INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)',
            (key, value)
        )
    
    conn.commit()
    conn.close()
    print("✓ Database seeded with default data")


# Helper functions for the web app

def get_all_skills(include_blacklisted=False):
    """Get all skills grouped by tier"""
    conn = get_db()
    cursor = conn.cursor()
    
    if include_blacklisted:
        cursor.execute('SELECT * FROM skills ORDER BY tier, name')
    else:
        cursor.execute('SELECT * FROM skills WHERE is_blacklisted = 0 ORDER BY tier, name')
    
    skills = cursor.fetchall()
    conn.close()
    
    # Group by tier
    grouped = {}
    for skill in skills:
        tier = skill['tier']
        if tier not in grouped:
            grouped[tier] = []
        grouped[tier].append(dict(skill))
    
    return grouped


def get_approved_skills():
    """Get list of approved (non-blacklisted) skill names"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT name FROM skills WHERE is_blacklisted = 0')
    skills = [row['name'] for row in cursor.fetchall()]
    conn.close()
    return skills


def get_blacklisted_skills():
    """Get list of blacklisted skill names"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT name FROM skills WHERE is_blacklisted = 1')
    skills = [row['name'] for row in cursor.fetchall()]
    conn.close()
    return skills


def get_skills_by_tier():
    """Get skills organized by tier (for CV generation)"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT name, tier FROM skills WHERE is_blacklisted = 0')
    
    tiers = {
        'tier_1_core': [],
        'tier_2_major': [],
        'tier_3_specialist': [],
        'tier_4_tools': [],
    }
    
    for row in cursor.fetchall():
        if row['tier'] in tiers:
            tiers[row['tier']].append(row['name'])
    
    conn.close()
    return tiers


def get_experiences():
    """Get all experiences with their bullet points"""
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM experiences ORDER BY order_index')
    experiences = []
    
    for exp in cursor.fetchall():
        exp_dict = dict(exp)
        
        # Get bullet points
        cursor.execute(
            'SELECT * FROM bullet_points WHERE experience_id = ? ORDER BY order_index',
            (exp['id'],)
        )
        exp_dict['bullet_points'] = [dict(bp) for bp in cursor.fetchall()]
        
        experiences.append(exp_dict)
    
    conn.close()
    return experiences


def get_settings():
    """Get all settings as a dictionary"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT key, value FROM settings')
    settings = {row['key']: row['value'] for row in cursor.fetchall()}
    conn.close()
    return settings


def update_setting(key, value):
    """Update a single setting"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute(
        'INSERT OR REPLACE INTO settings (key, value, updated_at) VALUES (?, ?, ?)',
        (key, value, datetime.now())
    )
    conn.commit()
    conn.close()


def add_skill(name, tier='tier_4_tools', category=None, is_blacklisted=False):
    """Add a new skill"""
    conn = get_db()
    cursor = conn.cursor()
    try:
        cursor.execute(
            'INSERT INTO skills (name, tier, category, is_blacklisted) VALUES (?, ?, ?, ?)',
            (name, tier, category, is_blacklisted)
        )
        conn.commit()
        skill_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        skill_id = None
    conn.close()
    return skill_id


def update_skill(skill_id, name=None, tier=None, is_blacklisted=None):
    """Update a skill"""
    conn = get_db()
    cursor = conn.cursor()
    
    updates = []
    values = []
    
    if name is not None:
        updates.append('name = ?')
        values.append(name)
    if tier is not None:
        updates.append('tier = ?')
        values.append(tier)
    if is_blacklisted is not None:
        updates.append('is_blacklisted = ?')
        values.append(is_blacklisted)
    
    if updates:
        values.append(skill_id)
        cursor.execute(f'UPDATE skills SET {", ".join(updates)} WHERE id = ?', values)
        conn.commit()
    
    conn.close()


def delete_skill(skill_id):
    """Delete a skill"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM skills WHERE id = ?', (skill_id,))
    conn.commit()
    conn.close()


def add_experience(company, role, start_date=None, end_date=None, is_current=False):
    """Add a new experience"""
    conn = get_db()
    cursor = conn.cursor()
    
    # Get max order index
    cursor.execute('SELECT MAX(order_index) FROM experiences')
    max_order = cursor.fetchone()[0] or 0
    
    cursor.execute('''
        INSERT INTO experiences (company, role, start_date, end_date, is_current, order_index)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (company, role, start_date, end_date, is_current, max_order + 1))
    
    exp_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return exp_id


def update_experience(exp_id, **kwargs):
    """Update an experience"""
    conn = get_db()
    cursor = conn.cursor()
    
    allowed = ['company', 'role', 'start_date', 'end_date', 'is_current', 'order_index']
    updates = []
    values = []
    
    for key, value in kwargs.items():
        if key in allowed:
            updates.append(f'{key} = ?')
            values.append(value)
    
    if updates:
        values.append(exp_id)
        cursor.execute(f'UPDATE experiences SET {", ".join(updates)} WHERE id = ?', values)
        conn.commit()
    
    conn.close()


def add_bullet_point(experience_id, base_description, tech_placeholder=None):
    """Add a bullet point to an experience"""
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT MAX(order_index) FROM bullet_points WHERE experience_id = ?', (experience_id,))
    max_order = cursor.fetchone()[0] or 0
    
    cursor.execute('''
        INSERT INTO bullet_points (experience_id, base_description, tech_placeholder, order_index)
        VALUES (?, ?, ?, ?)
    ''', (experience_id, base_description, tech_placeholder, max_order + 1))
    
    bp_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return bp_id


def update_bullet_point(bp_id, base_description=None, tech_placeholder=None):
    """Update a bullet point"""
    conn = get_db()
    cursor = conn.cursor()
    
    updates = []
    values = []
    
    if base_description is not None:
        updates.append('base_description = ?')
        values.append(base_description)
    if tech_placeholder is not None:
        updates.append('tech_placeholder = ?')
        values.append(tech_placeholder)
    
    if updates:
        values.append(bp_id)
        cursor.execute(f'UPDATE bullet_points SET {", ".join(updates)} WHERE id = ?', values)
        conn.commit()
    
    conn.close()


def delete_bullet_point(bp_id):
    """Delete a bullet point"""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM bullet_points WHERE id = ?', (bp_id,))
    conn.commit()
    conn.close()


# Initialize on import
init_db()
seed_default_data()
