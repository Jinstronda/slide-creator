# API Testing Guide

## Start the API Server

```powershell
cd "C:\Users\joaop\Documents\Augusta Labs\Cases Study Slide Optimization\case_study_generator"
uv run python run_api.py
```

The API will be available at: `http://localhost:8000`

**API Documentation (Swagger UI)**: `http://localhost:8000/docs`

## Test Endpoints

### 1. Health Check

```powershell
curl http://localhost:8000/health
```

Expected response:
```json
{"status":"healthy","service":"case-study-generator"}
```

### 2. Generate Presentation (1 Case)

```powershell
curl -X POST http://localhost:8000/api/generate `
  -H "Content-Type: application/json" `
  -d '{
    "company_name": "MedTech Solutions",
    "company_description": "Healthcare SaaS company with 200 employees focused on patient engagement and clinical workflow automation in hospitals and clinics across Europe.",
    "presentation_type": 1
  }' `
  --output test_1case.pptx
```

### 3. Generate Presentation (2 Cases)

```powershell
curl -X POST http://localhost:8000/api/generate `
  -H "Content-Type: application/json" `
  -d '{
    "company_name": "FinTech Innovations",
    "company_description": "Financial technology startup specializing in payment processing and fraud detection using machine learning for e-commerce platforms.",
    "presentation_type": 2
  }' `
  --output test_2cases.pptx
```

### 4. Generate Presentation (4 Cases)

```powershell
curl -X POST http://localhost:8000/api/generate `
  -H "Content-Type: application/json" `
  -d '{
    "company_name": "Global Logistics Corp",
    "company_description": "Large logistics and supply chain management company operating warehouses and distribution centers across 15 countries, focusing on automation and real-time tracking.",
    "presentation_type": 4
  }' `
  --output test_4cases.pptx
```

## Test Error Cases

### Invalid Presentation Type

```powershell
curl -X POST http://localhost:8000/api/generate `
  -H "Content-Type: application/json" `
  -d '{
    "company_name": "Test Company",
    "company_description": "A test company description that is long enough to pass validation.",
    "presentation_type": 3
  }'
```

Expected: 422 Unprocessable Entity (validation error)

### Missing API Key

Temporarily rename `.env` file and restart server, then:

```powershell
curl -X POST http://localhost:8000/api/generate `
  -H "Content-Type: application/json" `
  -d '{
    "company_name": "Test Company",
    "company_description": "A test company description.",
    "presentation_type": 4
  }'
```

Expected: 401 Unauthorized

### Short Description

```powershell
curl -X POST http://localhost:8000/api/generate `
  -H "Content-Type: application/json" `
  -d '{
    "company_name": "Test",
    "company_description": "Short",
    "presentation_type": 4
  }'
```

Expected: 422 Unprocessable Entity (description too short)

## Performance Testing

Typical response times:
- **AI Selection**: 2-5 seconds
- **PPTX Generation**: 1-3 seconds
- **Total**: 3-8 seconds per request

## Notes

- All generated PPTX files will be downloaded to the current directory
- Filenames follow the pattern: `{company_name}_{N}-cases_{timestamp}.pptx`
- API uses streaming response (no temp files created)
- CORS is enabled for all origins (update `api/app.py` for production)
