"""Configuration for template placeholders and mappings."""

# File paths - Change these to use different templates/data
DEFAULT_TEMPLATE = "templates/Case studies Template (1).pptx"
DEFAULT_DATA_FILE = "data/case_studies_complete.json"
DEFAULT_OUTPUT_DIR = "output"

# Template placeholder configuration
TEMPLATE_CONFIG = {
    # Slide-level placeholders
    "slide_title": "slide_case_studies_title",  # Company Name - Selected Case Studies
    "slide_subtitle": "slide_case_studies_subtitle",
    "4_cases_title": "4_cases_title",
    "slide_number": "sn",  # Page number for each slide
    
    # Per-case-study placeholders (will be formatted with case number)
    "case_study_name": "case_study_{n}_name",
    "case_study_title": "case_study_{n}_title",
    "case_study_description": "case_study_{n}_description",
    "case_study_image": "case_study_{n}_image",
    "case_study_metric": "n{n}",  # n1, n2, n3, n4
    "metric_label": "metric_label_case_study_{n}",
    "case_study_category": "case_study_{n}_category",
    "tab_label": "tab_{n}_label",
    
    # Slide 2: Challenge/Solution/Impact for all 4 case studies
    "case_study_challenge": "case_study_{n}_challenge_{x}",  # n=1-4, x=1-3
    "case_study_solution": "case_study_{n}_solution_{x}",  # n=1-4, x=1-4
    "case_study_impact": "case_study_{n}_impact_{x}",  # n=1-4, x=1-3
    
    # Slide 3: Single detailed case study (case study 1 only - unnumbered)
    "case_study_name_slide3": "case_study_name",  # Name without number
    "case_study_category_slide3": "case_study_category",  # Category without number
    "metric_label_slide3": "metric_label_case_study",  # Metric label without number
    "challenge": "challenge_{x}",  # x = 1-3
    "solution_intro": "solution_intro",
    "solution": "solution_{x}",  # x = 1-4
    "impact": "impact_{x}",  # x = 1-3
    
    # Company metadata
    "company_name": "company_name",
    "company_description": "company_description",
    "generation_date": "generation_date"
}

# Default values
DEFAULT_SUBTITLE = "Selected Case Studies"
DEFAULT_4_CASES_TITLE = "4 Selected Case Studies"
DEFAULT_IMAGE_PLACEHOLDER = "[Image placeholder]"
DEFAULT_METRIC_PLACEHOLDER = "-"
DEFAULT_CATEGORY = ""

# Length limits
MAX_METRIC_LENGTH = 40
MAX_TITLE_LENGTH = 60

# Industry category mapping
INDUSTRY_CATEGORIES = {
    "infrastructure": ["Ferrovial", "Quadrante", "Nortecnica"],
    "logistics": ["Portir", "AddVolt", "Profit"],
    "media": ["24 horas", "Media Capital", "SportTV"],
    "retail": ["Farmácias", "STS", "Rádio Popular", "Bene"],
    "healthcare": ["Astrazeneca", "Maia", "Sword Health", "Vet AI", "Medicare", "Pet24"],
    "financial": ["Banco", "BIG", "Lince Capital", "Sage", "Millennium"],
    "technology": ["Clever", "DeskSkill", "Gist", "Code for All"],
    "public sector": ["SPMS", "Câmara", "FAP", "SST"],
}

def get_industry_category(org_name: str) -> str:
    """Determine industry category based on organization name."""
    org_lower = org_name.lower()
    for category, keywords in INDUSTRY_CATEGORIES.items():
        for keyword in keywords:
            if keyword.lower() in org_lower:
                return category.upper()
    return "TECHNOLOGY"  # Default category

