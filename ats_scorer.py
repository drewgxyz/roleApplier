"""
ATS (Applicant Tracking System) Scoring Module

Based on research from:
- Jobscan ATS research (2023-2024)
- TopResume ATS studies
- LinkedIn Talent Solutions reports
- HR technology industry standards

Key ATS factors:
1. Keyword matching (exact + semantic)
2. Section structure and headers
3. Formatting simplicity
4. Contact information completeness
5. Skills alignment
6. Experience relevance
7. Education presence
8. Quantified achievements
9. File format compatibility
10. Length appropriateness
"""

import re
from collections import Counter
from typing import Dict, List, Tuple


class ATSScorer:
    """
    ATS Compatibility Scorer
    
    Scoring methodology based on industry research:
    - 40% Keyword Match (most critical for ATS parsing)
    - 20% Skills Alignment
    - 15% Structure & Formatting
    - 10% Quantified Achievements
    - 10% Experience Relevance
    - 5% Contact Info Completeness
    """
    
    # Standard ATS-friendly section headers
    STANDARD_SECTIONS = [
        'summary', 'professional summary', 'profile', 'objective',
        'experience', 'work experience', 'professional experience', 'employment',
        'education', 'academic background',
        'skills', 'technical skills', 'core competencies', 'expertise',
        'certifications', 'certificates',
        'projects', 'key projects',
    ]
    
    # Characters that can cause ATS parsing issues
    PROBLEMATIC_CHARS = ['│', '║', '┃', '▪', '▸', '►', '★', '●', '○', '◆', '◇', '→', '⟶', '✓', '✔', '✗', '✘']
    
    def __init__(self):
        self.scores = {}
        self.feedback = []
        self.warnings = []
    
    def score_cv(self, cv_text: str, job_description: str, job_skills: List[str]) -> Dict:
        """
        Score a CV against a job description
        Returns detailed scoring breakdown
        """
        self.scores = {}
        self.feedback = []
        self.warnings = []
        
        # 1. Keyword Match Score (40%)
        keyword_score, keyword_details = self._score_keywords(cv_text, job_description, job_skills)
        self.scores['keywords'] = {'score': keyword_score, 'weight': 0.40, 'details': keyword_details}
        
        # 2. Skills Alignment (20%)
        skills_score, skills_details = self._score_skills_alignment(cv_text, job_skills)
        self.scores['skills'] = {'score': skills_score, 'weight': 0.20, 'details': skills_details}
        
        # 3. Structure & Formatting (15%)
        structure_score, structure_details = self._score_structure(cv_text)
        self.scores['structure'] = {'score': structure_score, 'weight': 0.15, 'details': structure_details}
        
        # 4. Quantified Achievements (10%)
        quant_score, quant_details = self._score_quantified_achievements(cv_text)
        self.scores['achievements'] = {'score': quant_score, 'weight': 0.10, 'details': quant_details}
        
        # 5. Experience Relevance (10%)
        exp_score, exp_details = self._score_experience_relevance(cv_text, job_description)
        self.scores['experience'] = {'score': exp_score, 'weight': 0.10, 'details': exp_details}
        
        # 6. Contact Info (5%)
        contact_score, contact_details = self._score_contact_info(cv_text)
        self.scores['contact'] = {'score': contact_score, 'weight': 0.05, 'details': contact_details}
        
        # Calculate weighted total
        total_score = sum(
            self.scores[key]['score'] * self.scores[key]['weight'] 
            for key in self.scores
        )
        
        # Generate grade
        grade = self._get_grade(total_score)
        
        return {
            'total_score': round(total_score, 1),
            'grade': grade,
            'breakdown': self.scores,
            'feedback': self.feedback,
            'warnings': self.warnings,
            'ats_pass_likelihood': self._get_pass_likelihood(total_score)
        }
    
    def _score_keywords(self, cv_text: str, job_description: str, job_skills: List[str]) -> Tuple[float, dict]:
        """Score keyword matching between CV and job description"""
        cv_lower = cv_text.lower()
        job_lower = job_description.lower()
        
        # Extract important keywords from job description
        # Focus on: skills, tools, technologies, action verbs
        job_words = re.findall(r'\b[a-zA-Z]{3,}\b', job_lower)
        job_word_freq = Counter(job_words)
        
        # Filter to meaningful keywords (appear 2+ times or are skills)
        important_keywords = set()
        for word, count in job_word_freq.items():
            if count >= 2 and word not in self._get_stop_words():
                important_keywords.add(word)
        
        # Add all job skills
        for skill in job_skills:
            important_keywords.add(skill.lower())
        
        # Count matches
        matched = []
        missing = []
        
        for keyword in important_keywords:
            if keyword in cv_lower or keyword.replace('-', ' ') in cv_lower:
                matched.append(keyword)
            else:
                missing.append(keyword)
        
        if not important_keywords:
            return 70.0, {'matched': 0, 'total': 0, 'missing': []}
        
        match_rate = len(matched) / len(important_keywords) * 100
        
        # Feedback
        if match_rate < 60:
            self.feedback.append(f"Low keyword match ({len(matched)}/{len(important_keywords)}). Consider adding: {', '.join(missing[:5])}")
        elif match_rate < 80:
            self.feedback.append(f"Good keyword coverage. Missing: {', '.join(missing[:3])}")
        
        return min(match_rate, 100), {
            'matched': len(matched),
            'total': len(important_keywords),
            'match_rate': f"{match_rate:.0f}%",
            'missing_top5': missing[:5]
        }
    
    def _score_skills_alignment(self, cv_text: str, job_skills: List[str]) -> Tuple[float, dict]:
        """Score how well CV skills align with job requirements"""
        cv_lower = cv_text.lower()
        
        matched_skills = []
        missing_skills = []
        
        for skill in job_skills:
            skill_lower = skill.lower()
            # Check for exact match or common variations
            if (skill_lower in cv_lower or 
                skill_lower.replace(' ', '-') in cv_lower or
                skill_lower.replace('-', ' ') in cv_lower):
                matched_skills.append(skill)
            else:
                missing_skills.append(skill)
        
        if not job_skills:
            return 80.0, {'matched': 0, 'required': 0}
        
        match_rate = len(matched_skills) / len(job_skills) * 100
        
        if match_rate < 50:
            self.warnings.append("Less than 50% of required skills mentioned")
        
        return min(match_rate, 100), {
            'matched': len(matched_skills),
            'required': len(job_skills),
            'matched_skills': matched_skills[:10],
            'missing_skills': missing_skills[:5]
        }
    
    def _score_structure(self, cv_text: str) -> Tuple[float, dict]:
        """Score CV structure and ATS-friendly formatting"""
        score = 100
        issues = []
        
        cv_lower = cv_text.lower()
        
        # Check for standard section headers
        sections_found = []
        for section in self.STANDARD_SECTIONS:
            if section in cv_lower:
                sections_found.append(section)
        
        if len(sections_found) < 3:
            score -= 20
            issues.append("Missing standard section headers")
        
        # Check for problematic characters
        for char in self.PROBLEMATIC_CHARS:
            if char in cv_text:
                score -= 5
                issues.append(f"Contains special character '{char}' that may confuse ATS")
                break
        
        # Check for tables (indicated by multiple tabs or complex spacing)
        if cv_text.count('\t\t') > 3:
            score -= 10
            issues.append("Complex tabular formatting detected")
        
        # Check length (word count)
        word_count = len(cv_text.split())
        if word_count < 200:
            score -= 15
            issues.append("CV appears too short")
        elif word_count > 800:
            score -= 10
            issues.append("CV may be too long for single page")
        
        if issues:
            self.feedback.extend(issues[:2])
        
        return max(score, 0), {
            'sections_found': sections_found,
            'word_count': word_count,
            'issues': issues
        }
    
    def _score_quantified_achievements(self, cv_text: str) -> Tuple[float, dict]:
        """Score presence of quantified achievements (numbers, percentages, metrics)"""
        
        # Patterns for quantified achievements
        patterns = [
            r'\d+%',  # Percentages
            r'\$[\d,]+',  # Dollar amounts
            r'\d+\+?\s*(years?|months?)',  # Time periods
            r'\d+\s*(team|people|members|engineers|developers)',  # Team sizes
            r'(reduced|increased|improved|grew|saved)\s+.*?\d+',  # Impact metrics
            r'\d+x',  # Multipliers
            r'\d{1,3}(,\d{3})+',  # Large numbers
        ]
        
        matches = []
        for pattern in patterns:
            found = re.findall(pattern, cv_text.lower())
            matches.extend(found)
        
        # Score based on number of quantified achievements
        num_metrics = len(matches)
        
        if num_metrics >= 8:
            score = 100
        elif num_metrics >= 5:
            score = 85
        elif num_metrics >= 3:
            score = 70
        elif num_metrics >= 1:
            score = 50
        else:
            score = 20
            self.feedback.append("Add more quantified achievements (numbers, percentages, metrics)")
        
        return score, {
            'metrics_found': num_metrics,
            'examples': matches[:5]
        }
    
    def _score_experience_relevance(self, cv_text: str, job_description: str) -> Tuple[float, dict]:
        """Score relevance of experience to job description"""
        
        # Extract action verbs from job description
        job_verbs = self._extract_action_verbs(job_description)
        cv_verbs = self._extract_action_verbs(cv_text)
        
        # Check overlap
        matching_verbs = set(job_verbs) & set(cv_verbs)
        
        if not job_verbs:
            return 75.0, {'matching_verbs': 0}
        
        verb_match_rate = len(matching_verbs) / len(set(job_verbs)) * 100
        
        # Also check for industry/domain keywords
        domain_keywords = ['software', 'engineer', 'develop', 'build', 'design', 'implement', 
                          'deploy', 'maintain', 'optimize', 'scale', 'integrate', 'automate']
        
        domain_matches = sum(1 for kw in domain_keywords if kw in cv_text.lower())
        domain_score = min(domain_matches / 6 * 100, 100)
        
        final_score = (verb_match_rate * 0.6) + (domain_score * 0.4)
        
        return min(final_score, 100), {
            'matching_action_verbs': list(matching_verbs)[:5],
            'domain_keyword_matches': domain_matches
        }
    
    def _score_contact_info(self, cv_text: str) -> Tuple[float, dict]:
        """Score completeness of contact information"""
        score = 0
        found = []
        
        # Email
        if re.search(r'[\w\.-]+@[\w\.-]+\.\w+', cv_text):
            score += 30
            found.append('email')
        
        # Phone
        if re.search(r'[\d\s\-\(\)]{10,}', cv_text):
            score += 25
            found.append('phone')
        
        # LinkedIn
        if 'linkedin' in cv_text.lower():
            score += 25
            found.append('linkedin')
        
        # Location
        location_patterns = ['london', 'uk', 'united kingdom', 'remote']
        if any(loc in cv_text.lower() for loc in location_patterns):
            score += 20
            found.append('location')
        
        if score < 80:
            self.feedback.append("Ensure contact info includes: email, phone, LinkedIn, location")
        
        return score, {'found': found}
    
    def _extract_action_verbs(self, text: str) -> List[str]:
        """Extract action verbs commonly used in CVs"""
        action_verbs = [
            'led', 'managed', 'developed', 'created', 'designed', 'implemented',
            'built', 'delivered', 'achieved', 'improved', 'increased', 'reduced',
            'optimized', 'automated', 'streamlined', 'coordinated', 'collaborated',
            'analyzed', 'architected', 'deployed', 'maintained', 'scaled',
            'integrated', 'migrated', 'refactored', 'mentored', 'trained'
        ]
        
        text_lower = text.lower()
        found = [verb for verb in action_verbs if verb in text_lower]
        return found
    
    def _get_stop_words(self) -> set:
        """Common words to ignore in keyword matching"""
        return {
            'the', 'and', 'for', 'with', 'you', 'your', 'our', 'will', 'are',
            'have', 'has', 'been', 'being', 'was', 'were', 'this', 'that',
            'from', 'they', 'their', 'what', 'which', 'who', 'whom', 'when',
            'where', 'why', 'how', 'all', 'each', 'every', 'both', 'few',
            'more', 'most', 'other', 'some', 'such', 'than', 'too', 'very',
            'can', 'could', 'may', 'might', 'must', 'shall', 'should', 'would',
            'about', 'above', 'after', 'again', 'against', 'into', 'through',
            'during', 'before', 'between', 'under', 'over', 'once', 'here',
            'there', 'these', 'those', 'then', 'just', 'also', 'only', 'own',
            'same', 'able', 'work', 'working', 'role', 'team', 'company',
            'experience', 'looking', 'join', 'opportunity', 'position'
        }
    
    def _get_grade(self, score: float) -> str:
        """Convert score to letter grade"""
        if score >= 90:
            return 'A+'
        elif score >= 85:
            return 'A'
        elif score >= 80:
            return 'A-'
        elif score >= 75:
            return 'B+'
        elif score >= 70:
            return 'B'
        elif score >= 65:
            return 'B-'
        elif score >= 60:
            return 'C+'
        elif score >= 55:
            return 'C'
        elif score >= 50:
            return 'C-'
        else:
            return 'D'
    
    def _get_pass_likelihood(self, score: float) -> str:
        """Estimate likelihood of passing ATS screening"""
        if score >= 85:
            return "Very High (90%+)"
        elif score >= 75:
            return "High (75-90%)"
        elif score >= 65:
            return "Moderate (50-75%)"
        elif score >= 50:
            return "Low (25-50%)"
        else:
            return "Very Low (<25%)"


def estimate_page_count(text: str, chars_per_page: int = 3000) -> float:
    """
    Estimate page count based on character count.
    Standard single-column CV: ~3000 characters per page
    """
    return len(text) / chars_per_page


def validate_single_page(cv_data: dict) -> Tuple[bool, float, List[str]]:
    """
    Validate that CV content fits on a single page.
    Returns: (is_valid, estimated_pages, suggestions)
    """
    suggestions = []
    
    # Estimate total character count
    total_chars = 0
    
    # Bio
    bio = cv_data.get('bio', '')
    total_chars += len(bio) * 1.2  # Bio has larger font weight
    
    # Expertise
    expertise = cv_data.get('expertise', [])
    total_chars += len(' • '.join(expertise)) * 1.1
    
    # Experience sections
    for section in ['t', 'a']:
        section_data = cv_data.get(section, {})
        skills = section_data.get('skills', '')
        total_chars += len(skills)
        
        for i in range(1, 5):
            bp = section_data.get(f'bp{i}', '')
            total_chars += len(bp) * 1.15  # Bullet formatting overhead
    
    # Add header/contact info estimate
    total_chars += 300
    
    # Estimate pages
    estimated_pages = total_chars / 2800  # Slightly conservative
    
    is_valid = estimated_pages <= 1.05  # Allow 5% overflow
    
    if not is_valid:
        overflow = (estimated_pages - 1) * 100
        suggestions.append(f"CV is ~{overflow:.0f}% over 1 page")
        
        # Specific suggestions
        if len(bio) > 400:
            suggestions.append("Shorten bio (currently too long)")
        if len(expertise) > 14:
            suggestions.append(f"Reduce expertise list from {len(expertise)} to 12-14")
        
        # Check bullet points
        for section in ['t', 'a']:
            section_data = cv_data.get(section, {})
            for i in range(1, 5):
                bp = section_data.get(f'bp{i}', '')
                if len(bp) > 350:
                    suggestions.append(f"Shorten {section}.bp{i} (currently {len(bp)} chars)")
    
    return is_valid, round(estimated_pages, 2), suggestions
