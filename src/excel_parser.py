"""Extract case study data from JSON file."""
import json
from typing import List, Dict, Any


def get_case_studies(json_path: str = None) -> List[Dict[str, Any]]:
    """Load and return case studies from JSON file."""
    if json_path is None:
        # Default to complete case studies file
        import os
        from pathlib import Path
        base_dir = Path(__file__).parent.parent
        json_path = os.path.join(base_dir, "data", "case_studies_complete.json")
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    return data['case_studies']
