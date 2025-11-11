"""FastAPI application for case study presentation generation."""
import os
from datetime import datetime
from pathlib import Path
from fastapi import FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from dotenv import load_dotenv

from .schemas import GenerateRequest, ErrorResponse

# Add parent directory to path for src imports
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.core import generate_presentation_to_memory

load_dotenv()

app = FastAPI(
    title="Case Study Presentation Generator API",
    description="AI-powered REST API for generating case study presentations",
    version="1.0.0"
)

# CORS configuration for frontend access
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Update with specific origins in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post(
    "/api/generate",
    response_class=StreamingResponse,
    responses={
        200: {"description": "PPTX file generated successfully"},
        400: {"model": ErrorResponse, "description": "Invalid request"},
        401: {"model": ErrorResponse, "description": "Missing API key"},
        500: {"model": ErrorResponse, "description": "Generation failed"}
    }
)
async def generate_presentation(request: GenerateRequest):
    """
    Generate case study presentation.

    Returns binary PPTX file with N case studies (1, 2, or 4).
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="OpenAI API key not configured"
        )

    base_dir = Path(__file__).parent.parent
    data_path = str(base_dir / "data" / "case_studies_complete.json")
    template_path = str(base_dir / "templates" / "Case studies Template (1).pptx")

    try:
        pptx_data = generate_presentation_to_memory(
            company_name=request.company_name,
            company_description=request.company_description,
            api_key=api_key,
            num_cases=request.presentation_type,
            data_path=data_path,
            template_path=template_path
        )

        safe_name = _sanitize_filename(request.company_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{safe_name}_{request.presentation_type}-cases_{timestamp}.pptx"

        return StreamingResponse(
            pptx_data,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'}
        )

    except FileNotFoundError as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Required file missing: {str(e)}"
        )
    except ValueError as e:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=str(e)
        )
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Presentation generation failed: {str(e)}"
        )


@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {"status": "healthy", "service": "case-study-generator"}


def _sanitize_filename(name: str) -> str:
    """Sanitize company name for safe filename."""
    import re
    safe = re.sub(r'[^\w\s-]', '', name)
    safe = re.sub(r'[-\s]+', '_', safe)
    return safe[:50]
