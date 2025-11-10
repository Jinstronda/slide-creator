"""CLI interface for case study generator."""
import os
import sys
import click
from dotenv import load_dotenv
from pathlib import Path

from .excel_parser import get_case_studies
from .ai_selector import select_case_studies, format_selected_for_pptx
from .pptx_generator import generate_presentation, add_company_context


# Import defaults from config
from .config import DEFAULT_TEMPLATE, DEFAULT_DATA_FILE, DEFAULT_OUTPUT_DIR

# Default paths
DEFAULT_DATA_PATH = DEFAULT_DATA_FILE
DEFAULT_TEMPLATE_PATH = DEFAULT_TEMPLATE


@click.command()
@click.option(
    "--company-name",
    required=True,
    help="Target company name"
)
@click.option(
    "--company-description",
    required=True,
    help="Detailed description of the company"
)
@click.option(
    "--output-dir",
    default=DEFAULT_OUTPUT_DIR,
    help="Output directory for generated presentations"
)
@click.option(
    "--template",
    default=DEFAULT_TEMPLATE_PATH,
    help="Path to PowerPoint template"
)
@click.option(
    "--data",
    default=DEFAULT_DATA_PATH,
    help="Path to Excel data file"
)
@click.option(
    "--api-key",
    help="OpenAI API key (or set OPENAI_API_KEY environment variable)"
)
def main(company_name, company_description, output_dir, template, data, api_key):
    """Generate a customized case study presentation."""
    load_dotenv()
    
    click.echo("Case Study Generator")
    click.echo("=" * 50)
    
    # Get API key
    api_key = api_key or os.getenv("OPENAI_API_KEY")
    if not api_key:
        click.echo("Error: OpenAI API key not found.", err=True)
        click.echo("Set OPENAI_API_KEY environment variable or use --api-key", err=True)
        sys.exit(1)
    
    # Resolve paths
    base_dir = Path(__file__).parent.parent
    data_path = _resolve_path(data, base_dir)
    template_path = _resolve_path(template, base_dir)
    output_path = _resolve_path(output_dir, base_dir)
    
    # Validate files exist
    if not os.path.exists(data_path):
        click.echo(f"Error: Data file not found: {data_path}", err=True)
        sys.exit(1)
    
    if not os.path.exists(template_path):
        click.echo(f"Error: Template file not found: {template_path}", err=True)
        sys.exit(1)
    
    try:
        # Step 1: Load case studies
        click.echo(f"\n[1/3] Loading case studies from {os.path.basename(data_path)}...")
        case_studies = get_case_studies(data_path)
        click.echo(f"      Found {len(case_studies)} case studies")
        
        # Step 2: AI selection
        click.echo(f"\n[2/3] Selecting best 4 case studies for {company_name}...")
        selected = select_case_studies(
            case_studies,
            company_name,
            company_description,
            api_key
        )
        click.echo("      Selected case studies:")
        for i, cs in enumerate(selected, 1):
            click.echo(f"      {i}. {cs['deal_title']} ({cs['org']})")
        
        # Step 3: Generate presentation
        click.echo(f"\n[3/3] Generating PowerPoint presentation...")
        placeholders = format_selected_for_pptx(selected)
        add_company_context(placeholders, company_name, company_description)
        
        # Debug: Show challenge/solution/impact placeholders
        click.echo("\n=== SLIDE 3 PLACEHOLDERS (Sample) ===")
        for key in sorted(placeholders.keys()):
            if 'challenge' in key or 'solution' in key or 'impact' in key:
                value = placeholders[key][:50] if placeholders[key] else "[empty]"
                click.echo(f"  {key}: {value}")
        
        output_file = generate_presentation(
            template_path,
            placeholders,
            output_path,
            company_name
        )
        
        click.echo(f"\nSUCCESS! Presentation generated:")
        click.echo(f"   {os.path.abspath(output_file)}")
        
    except Exception as e:
        click.echo(f"\nERROR: {str(e)}", err=True)
        import traceback
        traceback.print_exc()
        sys.exit(1)


def _resolve_path(path: str, base_dir: Path) -> str:
    """Resolve relative paths relative to base directory."""
    path_obj = Path(path)
    if path_obj.is_absolute():
        return str(path_obj)
    return str(base_dir / path_obj)


if __name__ == "__main__":
    main()

