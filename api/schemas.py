"""Pydantic schemas for API request/response validation."""
from pydantic import BaseModel, Field, field_validator
from typing import Literal


class GenerateRequest(BaseModel):
    """Request schema for presentation generation."""

    company_name: str = Field(
        ...,
        min_length=1,
        max_length=200,
        description="Target company name"
    )
    company_description: str = Field(
        ...,
        min_length=10,
        max_length=2000,
        description="Detailed company description"
    )
    presentation_type: Literal[0, 1, 2, 4] = Field(
        ...,
        description="Number of case studies: 0=all slides, 1=single case, 2=two cases, 4=four cases"
    )

    @field_validator('company_name')
    @classmethod
    def validate_name(cls, v: str) -> str:
        """Sanitize company name."""
        if not v.strip():
            raise ValueError("Company name cannot be empty")
        return v.strip()

    @field_validator('company_description')
    @classmethod
    def validate_description(cls, v: str) -> str:
        """Validate description length."""
        stripped = v.strip()
        if len(stripped) < 10:
            raise ValueError("Description too short (min 10 chars)")
        return stripped


class ErrorResponse(BaseModel):
    """Error response schema."""

    error: str = Field(..., description="Error message")
    detail: str | None = Field(None, description="Detailed error info")
