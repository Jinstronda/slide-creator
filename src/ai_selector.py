"""AI-powered case study selection."""
import json
import os
from pathlib import Path
from typing import List, Dict, Any
from openai import OpenAI
from .config import TEMPLATE_CONFIG, DEFAULT_SUBTITLE, DEFAULT_IMAGE_PLACEHOLDER, DEFAULT_METRIC_PLACEHOLDER


def select_case_studies(
    case_studies: List[Dict[str, Any]],
    company_name: str,
    company_description: str,
    api_key: str
) -> List[Dict[str, Any]]:
    """Select 4 most relevant case studies using AI."""
    # Filter to only case studies with Challenge/Solution/Impact data
    csi_cases = [cs for cs in case_studies if 'challenges' in cs and 'solutions' in cs and 'impacts' in cs]
    
    # If we don't have at least 4 with CSI data, fall back to all cases
    if len(csi_cases) < 4:
        print(f"Warning: Only {len(csi_cases)} case studies have Challenge/Solution/Impact data, using all cases")
        cases_to_use = case_studies
    else:
        print(f"Using {len(csi_cases)} case studies with Challenge/Solution/Impact data")
        cases_to_use = csi_cases
    
    client = OpenAI(api_key=api_key)
    
    prompt = _build_prompt(cases_to_use, company_name, company_description)
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an expert business analyst selecting relevant case studies."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"}
    )
    
    result = json.loads(response.choices[0].message.content)
    selected_indices = result.get("selected_indices", [])
    
    if len(selected_indices) != 4:
        raise ValueError(f"Expected 4 selections, got {len(selected_indices)}")
    
    return [cases_to_use[i] for i in selected_indices if i < len(cases_to_use)]


def _build_prompt(case_studies: List[Dict[str, Any]], company_name: str, company_description: str) -> str:
    """Build AI selection prompt."""
    case_list = "\n\n".join(
        f"Index: {i}\nTitle: {cs['deal_title']}\nOrg: {cs['org']}\nAngles: {', '.join(cs['angles'])}\nComments: {cs.get('comments', '')}"
        for i, cs in enumerate(case_studies)
    )
    
    return f"""Select the 4 most relevant case studies for this company:

TARGET COMPANY:
Name: {company_name}
Description: {company_description}

CRITERIA: Industry alignment, similar challenges, complementary perspectives, relevant use cases

CASE STUDIES:
{case_list}

Return JSON:
{{"reasoning": "why these 4", "selected_indices": [i1, i2, i3, i4]}}"""


def format_selected_for_pptx(selected: List[Dict[str, Any]]) -> Dict[str, str]:
    """Format case studies for PowerPoint placeholders using template config."""
    from .config import (TEMPLATE_CONFIG, DEFAULT_SUBTITLE, DEFAULT_IMAGE_PLACEHOLDER, 
                        DEFAULT_METRIC_PLACEHOLDER, MAX_TITLE_LENGTH, DEFAULT_4_CASES_TITLE)
    from .excel_parser import get_case_studies
    
    placeholders = {
        TEMPLATE_CONFIG["slide_subtitle"]: DEFAULT_SUBTITLE,
        TEMPLATE_CONFIG["4_cases_title"]: DEFAULT_4_CASES_TITLE
    }
    
    # Track which images we're using
    used_images = set()
    used_metrics = set()
    
    # First pass: assign images and metrics that exist (exclude dashes - will be replaced)
    for i, cs in enumerate(selected, 1):
        image_file = cs.get('image_file', '')
        if image_file:
            used_images.add(image_file)
        
        metric = cs.get('metric', '')
        # Only track real metrics (not dashes - those will be replaced via AI)
        if metric and metric not in ['—', '-', '']:
            used_metrics.add(metric)
    
    # Second pass: fill in missing images with AI matching
    for i, cs in enumerate(selected, 1):
        # Case study name (organization)
        name_key = TEMPLATE_CONFIG["case_study_name"].format(n=i)
        placeholders[name_key] = cs['org'].strip()
        
        # Case study title (from JSON)
        title_key = TEMPLATE_CONFIG["case_study_title"].format(n=i)
        title = cs.get('title', cs['deal_title']).strip()
        if len(title) > MAX_TITLE_LENGTH:
            title = title[:MAX_TITLE_LENGTH-3] + "..."
        placeholders[title_key] = title
        
        # Case study description (from JSON) - strip whitespace and clean prefixes
        desc_key = TEMPLATE_CONFIG["case_study_description"].format(n=i)
        raw_description = cs.get('description', '').strip()
        clean_description = _clean_description(raw_description)
        placeholders[desc_key] = clean_description
        
        # Image - use existing or find similar
        image_key = TEMPLATE_CONFIG["case_study_image"].format(n=i)
        image_file = cs.get('image_file', '')
        
        if image_file:
            # Has image - use it
            placeholders[image_key] = f"images/{image_file}"
        else:
            # No image - find similar company with image
            print(f"WARNING: No image for {cs['org']}, finding similar company...")
            similar_image = _find_similar_company_image(cs, get_case_studies(), used_images)
            if similar_image:
                placeholders[image_key] = f"images/{similar_image}"
                used_images.add(similar_image)
                print(f"OK: Using image from similar company: {similar_image}")
            else:
                placeholders[image_key] = DEFAULT_IMAGE_PLACEHOLDER
        
        # Metric (from JSON - e.g., "125", "#1", "-80%", "40K")
        metric_key = TEMPLATE_CONFIG["case_study_metric"].format(n=i)
        raw_metric = cs.get('metric', '')
        
        print(f"\n[Case {i}: {cs['org']}]")
        print(f"  Original metric: '{raw_metric}'")
        
        # If no metric or has dash/em-dash, find similar company's metric
        if not raw_metric or raw_metric in ['—', '-']:
            print(f"  WARNING: MISSING METRIC - finding similar company...")
            similar_metric, similar_label = _find_similar_company_metric(cs, get_case_studies(), used_metrics)
            if similar_metric:
                raw_metric = similar_metric
                used_metrics.add(similar_metric)
                print(f"  OK: Using metric from similar company: {similar_metric} ({similar_label})")
                # Also update metric label if found
                if similar_label and not cs.get('metric_label'):
                    cs['metric_label'] = similar_label
            else:
                raw_metric = DEFAULT_METRIC_PLACEHOLDER
                print(f"  ERROR: No similar metric found, using default: {DEFAULT_METRIC_PLACEHOLDER}")
        else:
            print(f"  OK: Has metric: {raw_metric}")
        
        # Add "+" to pure numbers that don't have special characters
        if raw_metric and raw_metric not in [DEFAULT_METRIC_PLACEHOLDER, "—", "-"]:
            # Check if it's a pure number (possibly with K/M suffix but no +, -, %, #)
            if raw_metric.replace('K', '').replace('M', '').replace('.', '').replace(',', '').isdigit():
                if not raw_metric.endswith('+'):
                    raw_metric = raw_metric + '+'
        
        placeholders[metric_key] = raw_metric
        
        # Metric label (from JSON - e.g., "Hours Saved per Month")
        metric_label_key = TEMPLATE_CONFIG["metric_label"].format(n=i)
        metric_label = cs.get('metric_label', '')
        # Normalize to consistent width (40 chars total) for alignment
        # Add 5 leading spaces, then pad to 45 total chars
        normalized_label = ("     " + metric_label).ljust(45)
        placeholders[metric_label_key] = normalized_label
        
        # Category (from JSON - e.g., "INFRASTRUCTURE", "MEDIA")
        category_key = TEMPLATE_CONFIG["case_study_category"].format(n=i)
        category = cs.get('category', 'TECHNOLOGY')
        placeholders[category_key] = category.upper()
        
        # Tab label (short org name)
        tab_key = TEMPLATE_CONFIG["tab_label"].format(n=i)
        placeholders[tab_key] = cs['org'].split()[0] if cs['org'] else f"Case {i}"
    
    # Third pass: Match logos to case studies using AI
    available_logos = _get_available_logos()
    used_logos = set()
    
    for i, cs in enumerate(selected, 1):
        logo_key = TEMPLATE_CONFIG["case_study_logo"].format(n=i)
        
        # Get API key from environment
        import os
        api_key = os.getenv("OPENAI_API_KEY")
        
        if available_logos and api_key:
            # Find available logos not yet used
            remaining_logos = [logo for logo in available_logos if logo not in used_logos]
            if remaining_logos:
                matched_logo = _match_logo_to_case_study(cs, remaining_logos, api_key)
                if matched_logo:
                    placeholders[logo_key] = f"Logos/{matched_logo}.svg"
                    used_logos.add(matched_logo)
                    print(f"  Matched logo: {matched_logo}")
                else:
                    placeholders[logo_key] = ""
            else:
                placeholders[logo_key] = ""
        else:
            placeholders[logo_key] = ""
        
        # Slide 2: Add Challenge/Solution/Impact for ALL case studies (numbered)
        if 'challenges' in cs and 'solutions' in cs and 'impacts' in cs:
            # Use pre-defined arrays from JSON
            challenge_points = cs['challenges'][:3]
            solution_points = cs['solutions'][:4]
            impact_points = cs['impacts'][:3]
        else:
            # Parse from description
            challenge_points, solution_points, impact_points = _parse_csi_description(cs.get('description', ''))
        
        # Add challenge bullet points for this case study (slide 2)
        for x in range(1, 4):
            challenge_key = TEMPLATE_CONFIG["case_study_challenge"].format(n=i, x=x)
            placeholders[challenge_key] = challenge_points[x-1] if x-1 < len(challenge_points) else ""
        
        # Add solution bullet points for this case study (slide 2)
        for x in range(1, 5):
            solution_key = TEMPLATE_CONFIG["case_study_solution"].format(n=i, x=x)
            placeholders[solution_key] = solution_points[x-1] if x-1 < len(solution_points) else ""
        
        # Add impact bullet points for this case study (slide 2)
        for x in range(1, 4):
            impact_key = TEMPLATE_CONFIG["case_study_impact"].format(n=i, x=x)
            placeholders[impact_key] = impact_points[x-1] if x-1 < len(impact_points) else ""
        
        # Slide 3: ALSO add simple names for FIRST case study only
        if i == 1:
            # Add name, category, and metric label for slide 3 (without numbers)
            placeholders[TEMPLATE_CONFIG["case_study_name_slide3"]] = cs['org'].strip()
            placeholders[TEMPLATE_CONFIG["case_study_category_slide3"]] = cs.get('category', 'TECHNOLOGY').upper()
            placeholders[TEMPLATE_CONFIG["metric_label_slide3"]] = cs.get('metric_label', '')
            
            # Add challenge bullet points (simple names for slide 3)
            for x in range(1, 4):
                challenge_key = TEMPLATE_CONFIG["challenge"].format(x=x)
                placeholders[challenge_key] = challenge_points[x-1] if x-1 < len(challenge_points) else ""
            
            # Add solution intro
            placeholders[TEMPLATE_CONFIG["solution_intro"]] = "Our Solution:"
            
            # Add solution bullet points (simple names for slide 3)
            for x in range(1, 5):
                solution_key = TEMPLATE_CONFIG["solution"].format(x=x)
                placeholders[solution_key] = solution_points[x-1] if x-1 < len(solution_points) else ""
            
            # Add impact bullet points (simple names for slide 3)
            for x in range(1, 4):
                impact_key = TEMPLATE_CONFIG["impact"].format(x=x)
                placeholders[impact_key] = impact_points[x-1] if x-1 < len(impact_points) else ""
    
    return placeholders


def _clean_description(description: str) -> str:
    """Remove Challenge/Solution/Impact prefixes from description."""
    import re
    # Remove "Challenge:", "Solution:", "Impact:" labels
    cleaned = re.sub(r'(Challenge|Solution|Impact):\s*', '', description, flags=re.IGNORECASE)
    return cleaned.strip()


def _parse_csi_description(description: str) -> tuple:
    """Parse Challenge/Solution/Impact sections from description."""
    import re
    
    # Extract sections
    challenge_match = re.search(r'Challenge:\s*(.+?)(?:Solution:|$)', description, re.IGNORECASE | re.DOTALL)
    solution_match = re.search(r'Solution:\s*(.+?)(?:Impact:|$)', description, re.IGNORECASE | re.DOTALL)
    impact_match = re.search(r'Impact:\s*(.+?)$', description, re.IGNORECASE | re.DOTALL)
    
    challenge_text = challenge_match.group(1).strip() if challenge_match else ""
    solution_text = solution_match.group(1).strip() if solution_match else ""
    impact_text = impact_match.group(1).strip() if impact_match else ""
    
    # Split into bullet points
    challenge_points = _split_into_bullets(challenge_text, max_bullets=3)
    solution_points = _split_into_bullets(solution_text, max_bullets=4)  # Support 4 solutions
    impact_points = _split_into_bullets(impact_text, max_bullets=3)
    
    return challenge_points, solution_points, impact_points


def _split_into_bullets(text: str, max_bullets: int = 3) -> list:
    """Split text into bullet points."""
    import re
    
    if not text:
        return []
    
    # Split by period or semicolon
    parts = re.split(r'[.;]\s+', text)
    
    # Clean and filter
    bullets = [p.strip() for p in parts if p.strip() and len(p.strip()) > 5]
    
    # Return up to max_bullets
    return bullets[:max_bullets]


def _find_similar_company_metric(target_case: Dict[str, Any], all_cases: List[Dict[str, Any]], 
                                  used_metrics: set) -> tuple:
    """Use AI to find the most similar company's metric that isn't being used."""
    import openai
    import os
    
    # Get case studies with metrics that we haven't used yet
    available_cases = [
        cs for cs in all_cases 
        if cs.get('metric') 
        and cs.get('metric') not in ['—', '-'] 
        and cs.get('metric') not in used_metrics
    ]
    
    if not available_cases:
        return None, None
    
    # Build prompt for AI
    target_desc = f"{target_case['org']} - {target_case.get('category', 'Unknown')} - {target_case.get('title', target_case['deal_title'])}"
    
    options = "\n".join([
        f"{i+1}. {cs['org']} - {cs.get('category', 'Unknown')} - {cs.get('title', cs['deal_title'])} (metric: {cs['metric']} - {cs.get('metric_label', '')})"
        for i, cs in enumerate(available_cases[:20])
    ])
    
    prompt = f"""Find the most similar company to borrow a metric from.

Target company (no metric available):
{target_desc}

Available companies with metrics:
{options}

Return ONLY the number (1-{min(20, len(available_cases))}) of the most similar company based on industry and solution type.

Return only the number."""

    try:
        client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        
        selection = response.choices[0].message.content.strip()
        selected_idx = int(selection) - 1
        
        if 0 <= selected_idx < len(available_cases):
            return available_cases[selected_idx]['metric'], available_cases[selected_idx].get('metric_label', '')
    except Exception as e:
        print(f"Warning: Could not find similar metric via AI: {e}")
    
    return None, None


def _find_similar_company_image(target_case: Dict[str, Any], all_cases: List[Dict[str, Any]], 
                                 used_images: set) -> str:
    """Use AI to find the most similar company that has an image."""
    import openai
    import os
    
    # Get case studies with images that we haven't used yet
    available_cases = [
        cs for cs in all_cases 
        if cs.get('image_file') and cs.get('image_file') not in used_images
    ]
    
    if not available_cases:
        return None
    
    # Build prompt for AI
    target_desc = f"{target_case['org']} - {target_case.get('category', 'Unknown')} - {target_case.get('title', target_case['deal_title'])}"
    
    options = "\n".join([
        f"{i+1}. {cs['org']} - {cs.get('category', 'Unknown')} - {cs.get('title', cs['deal_title'])} (image: {cs['image_file']})"
        for i, cs in enumerate(available_cases[:20])  # Limit to top 20 for token efficiency
    ])
    
    prompt = f"""You need to find the most similar company to use as a visual placeholder.

Target company (no image available):
{target_desc}

Available companies with images:
{options}

Return ONLY the number (1-{min(20, len(available_cases))}) of the most similar company based on:
- Industry/sector similarity
- Type of solution (AI, ML, automation, etc.)
- Company type (SaaS, infrastructure, etc.)

Return only the number, nothing else."""

    try:
        client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        
        selection = response.choices[0].message.content.strip()
        selected_idx = int(selection) - 1
        
        if 0 <= selected_idx < len(available_cases):
            return available_cases[selected_idx]['image_file']
    except Exception as e:
        print(f"Warning: Could not find similar image via AI: {e}")
    
    return None


def _get_available_logos() -> List[str]:
    """Get list of available logo files from Logos directory."""
    project_root = Path(__file__).parent.parent
    logos_dir = project_root / "Logos"
    
    if not logos_dir.exists():
        return []
    
    # Get all SVG files
    logo_files = list(logos_dir.glob("*.svg"))
    # Extract label names from filenames (remove .svg extension)
    logo_labels = [f.stem for f in logo_files]
    
    return logo_labels


def _match_logo_to_case_study(case_study: Dict[str, Any], available_logos: List[str], api_key: str) -> str:
    """Use AI to match case study to best business value logo."""
    if not available_logos:
        return None
    
    client = OpenAI(api_key=api_key)
    
    # Build context from case study
    title = case_study.get('title', case_study.get('deal_title', ''))
    impacts = case_study.get('impacts', [])
    angles = case_study.get('angles', [])
    metric_label = case_study.get('metric_label', '')
    
    impact_text = "; ".join(impacts) if impacts else ""
    angles_text = "; ".join(angles) if angles else ""
    
    # Build prompt
    logos_list = "\n".join([f"{i+1}. {logo}" for i, logo in enumerate(available_logos)])
    
    prompt = f"""Match this case study to the BEST business value category.

Case Study:
Title: {title}
Impact: {impact_text}
Angles: {angles_text}
Metric: {metric_label}

Available Business Value Categories:
{logos_list}

Return ONLY the number (1-{len(available_logos)}) of the best matching category based on the primary business value/outcome delivered."""

    try:
        response = client.chat.completions.create(
            model="gpt-5-mini",
            max_completion_tokens=5000,
            messages=[{
                "role": "user",
                "content": prompt
            }]
        )
        
        selection = response.choices[0].message.content.strip()
        selected_idx = int(selection) - 1
        
        if 0 <= selected_idx < len(available_logos):
            return available_logos[selected_idx]
    except Exception as e:
        print(f"Warning: Could not match logo via AI: {e}")
    
    return None
