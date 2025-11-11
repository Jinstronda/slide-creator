"""CLI interface for case study generator."""
import os
import sys
import click
from dotenv import load_dotenv
from pathlib import Path

from .core import generate_presentation_workflow


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
        click.echo(f"\n[1/2] Generating presentation for {company_name}...")
        click.echo(f"      Using {os.path.basename(data_path)}")

        output_file = generate_presentation_workflow(
            company_name=company_name,
            company_description=company_description,
            api_key=api_key,
            num_cases=4,  # CLI defaults to 4 cases
            data_path=data_path,
            template_path=template_path,
            output_dir=output_path
        )

        click.echo(f"\n[2/2] SUCCESS! Presentation generated:")
        click.echo(f"      {os.path.abspath(output_file)}")

    except FileNotFoundError as e:
        click.echo(f"\nERROR: {str(e)}", err=True)
        sys.exit(1)
    except ValueError as e:
        click.echo(f"\nERROR: {str(e)}", err=True)
        sys.exit(1)
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

