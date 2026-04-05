# AI-Powered CV Customizer

Automatically customize your CV for any job posting using AI. Just paste the entire job description, and get a perfectly tailored CV in seconds.

## 🌐 Web Application (Recommended - Batch Mode)

**New!** Use the Flask web app to generate CVs and cover letters for multiple jobs at once.

### Quick Start (Web App)

1. **Set up virtual environment**:
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```

2. **Get OpenAI API key** from [OpenAI Platform](https://platform.openai.com/api-keys)

3. **Create `.env` file** in the project root:
   ```bash
   OPENAI_API_KEY=your_openai_api_key_here
   ```

4. **Run the web app**:
   ```bash
   source venv/bin/activate
   python web_app.py
   ```

5. **Open** http://localhost:5000 in your browser

### Features
- **Batch Processing**: Paste multiple job descriptions separated by `---` or triple newlines
- **Single LLM Call**: CV and cover letter generated together for efficiency
- **Modern UI**: Clean, responsive interface with TailwindCSS
- **Download All**: Get a ZIP file with all generated applications
- **Uses GPT-4o**: Latest OpenAI model for best results

---

## �️ CLI Mode (Original - Single Job)

For single job processing using Claude API:

### Quick Start (CLI)

1. **Install dependencies**:
   ```bash
   pip install anthropic python-docx reportlab
   ```

2. **Get Claude API key** from [Anthropic Console](https://console.anthropic.com/)

3. **Set up configuration**:
   - Run `python main.py` once to create `config/settings.json`
   - Add your API key to the config file

4. **Add your files**:
   - Place your CV template in `templates/template copy.docx`
   - Place your original CV in `data/orig.docx`

5. **Run the app**:
   ```bash
   python main.py
   ```

## 📁 File Structure

```
cv-customizer/
├── main.py                    # Main application
├── job_parser.py             # AI job description parser
├── experience_adapter.py     # AI experience adapter  
├── cv_customizer.py          # CV template engine
├── cover_letter_generator.py # AI cover letter generator
├── config/
│   └── settings.json         # Configuration (API key, paths, skills)
├── resources/
│   ├── template.docx         # Your CV template with placeholders
│   └── orig.docx            # Your original CV for reference
├── data/
│   ├── project_context.json # Additional project details (optional)
│   └── input.json           # Job description input file
└── outputs/                 # Generated CVs (auto-created)
    └── YYYYMMDD_HHMMSS-Company-Role/
        ├── Drew_Gillies_Software_Resume.docx
        ├── Drew_Gillies_Software_Resume.pdf
        ├── Drew_Gillies_Cover_Letter_Company.pdf
        └── Original_Job_Description.txt
```

## 🎯 How It Works

1. **AI parses job description** → Extracts requirements, skills, responsibilities
2. **AI adapts your experience** → Matches your projects to job requirements  
3. **Generates tailored CV** → Creates Word + PDF with perfect formatting
4. **Creates cover letter** → AI-generated, professional, tailored to role
5. **Smart optimization** → Automatically adjusts detail level for maximum impact

## ⚙️ Customizing Content Length and Detail

### 🔧 **Expertise Section Length**

The expertise section length is controlled by the relevance scoring system:

**Location**: `experience_adapter.py` → `_determine_content_strategy()` and `_get_enhanced_content_strategy()`

```python
# Standard Strategy (first pass)
'expertise_count': 14  # High relevance jobs
'expertise_count': 12  # Medium-high relevance  
'expertise_count': 10  # Medium relevance
'expertise_count': 8   # Low relevance

# Enhanced Strategy (when extra space available)  
'expertise_count': 16  # High relevance jobs (enhanced)
'expertise_count': 14  # Medium-high relevance (enhanced)
'expertise_count': 12  # Medium relevance (enhanced)
'expertise_count': 10  # Low relevance (enhanced)
```

**To increase expertise skills globally:**
1. Open `experience_adapter.py`
2. Find `_determine_content_strategy()` method
3. Increase `expertise_count` values by 2-4
4. Find `_get_enhanced_content_strategy()` method  
5. Increase those values by 2-4 as well

### 📝 **Bio Length Control**

**Location**: Same methods as above

```python
'bio_sentences': 3  # Shorter bio (more space for bullets)
'bio_sentences': 4  # Longer bio (less space for bullets)
```

**Bio length affects page balance:**
- **3 sentences**: More space for detailed bullet points
- **4 sentences**: More professional context, shorter bullets

### 🎯 **Bullet Point Detail Levels**

**Location**: `experience_adapter.py` → Detail level definitions

```python
# Current word count targets
'MAXIMUM': '40-55 words'    # Ultra-detailed for perfect job matches
'HIGH': '30-40 words'       # Detailed with metrics and technologies
'MEDIUM': '20-30 words'     # Balanced detail and impact
'LOW': '15-25 words'        # Concise but impactful
'MINIMAL': '10-15 words'    # Essential impact only
```

**To adjust bullet point length:**
1. Find `ENHANCED DETAIL LEVEL DEFINITIONS` in `regenerate_with_enhanced_detail()`
2. Increase word counts by 5-10 words per level
3. Also update the standard definitions in `adapt_experience_to_job()`

### 🏢 **Tech Stack Length**

**Location**: AI prompts in both methods

```python
# Current settings
"Tech stacks: Include MORE technologies (10-12)"  # Enhanced mode
"Tech stacks: Maximum 8-10 technologies"         # Standard mode
```

**To increase tech stack length:**
1. Search for "tech stack" in prompts
2. Change "8-10" to "10-12" and "10-12" to "12-15"

### 🎚️ **Relevance Scoring Thresholds**

**Location**: `experience_adapter.py` → `_determine_content_strategy()`

```python
if relevance_score >= 8:    # High detail threshold
elif relevance_score >= 6:  # Medium-high detail
elif relevance_score >= 4:  # Medium detail  
else:                       # Low detail
```

**To make more jobs get detailed treatment:**
- Lower thresholds: Change `>= 8` to `>= 7`, `>= 6` to `>= 5`, etc.
- This makes more jobs qualify for enhanced detail levels

### 📊 **Page Space Estimation**

**Location**: `cv_customizer.py` → `estimate_page_usage()`

```python
char_weights = {
    'bio': 2.8,           # Bio formatting weight
    'expertise': 1.2,     # Skills list weight  
    'tech_stack': 1.0,    # Single line weight
    'bullet_point': 2.2   # Bullet formatting weight
}

page_usage = total_chars / 3000.0  # Page capacity threshold
```

**To adjust when enhancement kicks in:**
1. Change the division number (3000) to be higher for earlier enhancement
2. Modify `if page_usage < 0.75` threshold in `main_app.py`

### 🔄 **Two-Pass Generation Control**

**Location**: `main_app.py` → `create_cv_from_input_file()`

```python
if page_usage < 0.75:  # Enhancement threshold (75% page usage)
```

**To change when enhancement happens:**
- `< 0.75` = Enhance if less than 75% page used
- `< 0.60` = Only enhance if less than 60% page used (more conservative)
- `< 0.85` = Enhance if less than 85% page used (more aggressive)

## 🎛️ **Quick Customization Examples**

### **Make Expertise Section Longer Globally**
```python
# In _determine_content_strategy() and _get_enhanced_content_strategy()
'expertise_count': 16  # Instead of 14 (high relevance)
'expertise_count': 14  # Instead of 12 (medium-high)
'expertise_count': 12  # Instead of 10 (medium)
```

### **Make Bullet Points More Detailed**
```python
# In ENHANCED DETAIL LEVEL DEFINITIONS
- MAXIMUM: 45-60 words  # Instead of 40-55
- HIGH: 35-45 words     # Instead of 30-40
- MEDIUM: 25-35 words   # Instead of 20-30
```

### **Make More Jobs Get Enhanced Detail**
```python
# Lower the thresholds
if relevance_score >= 6:    # Instead of >= 8
elif relevance_score >= 4:  # Instead of >= 6
elif relevance_score >= 2:  # Instead of >= 4
```

### **Make Enhancement More Aggressive**
```python
# In main_app.py
if page_usage < 0.85:  # Instead of < 0.75
```

## 🎯 **Understanding the Content Strategy**

The system uses a **two-pass approach**:

### **Pass 1: Adaptive Generation**
- Assesses job relevance (1-10 score)
- Generates CV with appropriate detail level
- Estimates page usage

### **Pass 2: Enhancement (if needed)**
- Only runs if page usage < 75%
- Uses enhanced detail levels
- Adds more technologies and metrics
- Maximizes impact within page limits

### **Relevance Scoring Factors**
1. **Core Technology Matches** (4 points max)
   - Python, Java, SQL, AWS mentions
2. **Domain Alignment** (4 points max)  
   - Finance, Cloud, Data keywords
3. **Seniority Match** (2 points max)
   - 2-3 years, mid-level experience

### **Detail Level Hierarchy**
```
Job Relevance → Detail Strategy → Content Length
10/10 → MAXIMUM → 40-55 word bullets, 16 skills
8/10  → HIGH    → 30-40 word bullets, 14 skills  
6/10  → MEDIUM  → 20-30 word bullets, 12 skills
4/10  → LOW     → 15-25 word bullets, 10 skills
```

## 📈 **Performance Optimization Tips**

### **Cost Management**
- Higher relevance thresholds = fewer enhanced generations = lower costs
- Most jobs use 1x API cost, perfect matches use 2x cost
- Average cost across all jobs: ~1.3x base cost

### **Quality vs Speed**
- Two-pass system maximizes quality for important applications
- Single-pass mode available by setting page threshold to 0.0
- Enhanced mode adds 30-60 seconds processing time

### **Page Length Control**
- System automatically prevents overflow beyond 1 page
- Character estimation is calibrated for your specific template
- Manual adjustment may be needed for template changes

## 🔧 **Advanced Configuration**

### **Skills Management**
Edit `config/settings.json`:

```json
{
  "skills_config": {
    "my_skills": {
      "tier_1_core": ["Python", "Java", "SQL", "AWS"],
      "tier_2_major": ["Docker", "PostgreSQL", "Linux"],
      "tier_3_specialist": ["Lambda", "SQS", "Redis"],
      "tier_4_tools": ["Git", "Jenkins", "pytest"]
    },
    "blacklisted_skills": ["Kubernetes", "React", "Angular"]
  }
}
```

### **Template Requirements**
Your `template copy.docx` needs these placeholders:
- `{{bio}}` - Professional summary
- `{{expertise}}` - Dynamic skills list
- `{{t.skills}}` - T. Rowe Price tech stack
- `{{t.bp1}}` through `{{t.bp4}}` - T. Rowe Price bullet points
- `{{a.skills}}` - AWS tech stack  
- `{{a.bp1}}` through `{{a.bp3}}` - AWS bullet points

### **Project Context Structure**
Create detailed context in `data/project_context.json`:

```json
{
  "t_rowe_price": {
    "bp1": {
      "detailed_description": "Extensive project details...",
      "technical_stack": ["Python", "PostgreSQL", "Docker"],
      "business_impact": "Reduced time by 85%",
      "team_size": "3 person team",
      "timeline": "4 months"
    }
  }
}
```

The AI uses this rich context to create more detailed, accurate bullet points.

## 🚨 **Troubleshooting Length Issues**

### **CV Too Long**
1. Reduce expertise_count by 2-4
2. Lower bullet point word counts by 5-10
3. Increase relevance score thresholds
4. Reduce bio to 3 sentences max

### **CV Too Short** 
1. Increase expertise_count by 2-4
2. Raise bullet point word counts by 5-10
3. Lower relevance score thresholds  
4. Reduce enhancement threshold to 0.85

### **Inconsistent Length**
1. Check page estimation calibration in `estimate_page_usage()`
2. Adjust character weights for your template
3. Test with different job types to validate

---

**The system automatically balances content length, detail level, and page constraints to give you the strongest possible CV for each role!** 🎯