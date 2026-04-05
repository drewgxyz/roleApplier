#!/usr/bin/env python3
"""
Process the ai_skills_review.txt file and add approved skills to the database.
Y = Add to approved skills (tier based on category)
N = Add to blacklist
? = Skip (don't add)
"""

from database import add_skill, get_db

# Mapping of categories to tiers
CATEGORY_TO_TIER = {
    'ai_framework': 'tier_2_major',
    'ai_api': 'tier_2_major',
    'ai_model': 'tier_3_specialist',
    'ai_tool': 'tier_3_specialist',
    'ai_concept': 'tier_3_specialist',
    'vector_db': 'tier_3_specialist',
    'ml_framework': 'tier_2_major',
    'ml_library': 'tier_3_specialist',
    'data_library': 'tier_3_specialist',
    'mlops': 'tier_3_specialist',
    'aws_ai': 'tier_3_specialist',
    'cloud_ai': 'tier_3_specialist',
    'messaging': 'tier_3_specialist',
    'orchestration': 'tier_3_specialist',
    'data_transform': 'tier_3_specialist',
    'data_warehouse': 'tier_3_specialist',
    'data_platform': 'tier_3_specialist',
    'data_processing': 'tier_3_specialist',
    'caching': 'tier_3_specialist',
    'task_queue': 'tier_4_tools',
    'framework': 'tier_2_major',
    'api': 'tier_3_specialist',
    'language': 'tier_2_major',
    'container_orchestration': 'tier_2_major',
    'gitops': 'tier_4_tools',
    'kubernetes': 'tier_4_tools',
    'iac': 'tier_2_major',
    'ci_cd': 'tier_3_specialist',
    'container': 'tier_3_specialist',
    'monitoring': 'tier_4_tools',
    'observability': 'tier_4_tools',
    'ai_observability': 'tier_4_tools',
    'testing': 'tier_4_tools',
    'load_testing': 'tier_4_tools',
    'ai_testing': 'tier_4_tools',
    'tool': 'tier_4_tools',
    'concept': 'tier_4_tools',
}

def process_review_file(filepath='ai_skills_review.txt'):
    added = []
    blacklisted = []
    skipped = []
    
    with open(filepath, 'r') as f:
        for line in f:
            line = line.strip()
            
            # Skip comments and empty lines
            if not line or line.startswith('#'):
                continue
            
            # Parse line: Y/N/? | Skill Name | Category
            parts = [p.strip() for p in line.split('|')]
            if len(parts) != 3:
                continue
            
            decision, skill_name, category = parts
            decision = decision.upper()
            
            if decision == 'Y':
                tier = CATEGORY_TO_TIER.get(category, 'tier_4_tools')
                result = add_skill(skill_name, tier=tier, category=category, is_blacklisted=False)
                if result:
                    added.append(f"{skill_name} ({tier})")
                else:
                    print(f"  ⚠ {skill_name} already exists")
            elif decision == 'N':
                result = add_skill(skill_name, tier='tier_4_tools', category=category, is_blacklisted=True)
                if result:
                    blacklisted.append(skill_name)
            else:
                skipped.append(skill_name)
    
    print("\n" + "="*50)
    print("SKILLS REVIEW COMPLETE")
    print("="*50)
    
    if added:
        print(f"\n✓ Added {len(added)} skills:")
        for s in added:
            print(f"  + {s}")
    
    if blacklisted:
        print(f"\n✗ Blacklisted {len(blacklisted)} skills:")
        for s in blacklisted:
            print(f"  - {s}")
    
    if skipped:
        print(f"\n? Skipped {len(skipped)} skills (marked with ?)")
    
    print("\nDone! Refresh the settings page to see changes.")

if __name__ == '__main__':
    process_review_file()
