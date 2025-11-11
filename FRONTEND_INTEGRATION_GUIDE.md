# Frontend Integration Guide

## API Overview

**Base URL (Production):** `https://your-app.up.railway.app`
**Base URL (Local Dev):** `http://localhost:8000`

**Single Endpoint:** `POST /api/generate`

---

## Request Format

### Endpoint
```
POST /api/generate
Content-Type: application/json
```

### Request Body Schema

```typescript
interface GenerateRequest {
  company_name: string;        // Required, 1-200 characters
  company_description: string; // Required, 10-2000 characters
  presentation_type: 1 | 2 | 4; // Required, only these values allowed
}
```

### Example Request Body

```json
{
  "company_name": "MedTech Solutions",
  "company_description": "Healthcare SaaS company with 200 employees focused on patient engagement and clinical workflow automation in hospitals across Europe.",
  "presentation_type": 4
}
```

### Validation Rules

| Field | Type | Required | Min Length | Max Length | Allowed Values |
|-------|------|----------|------------|------------|----------------|
| `company_name` | string | ✅ Yes | 1 | 200 | Any non-empty string |
| `company_description` | string | ✅ Yes | 10 | 2000 | Any descriptive text |
| `presentation_type` | integer | ✅ Yes | - | - | `0`, `1`, `2`, or `4` only |

**Important:**
- `presentation_type` controls which slides are returned:
  - `0` = All 3 slides (overview + grid + detailed single case)
  - `1` = Only Slide 3 (single detailed case study)
  - `2` = Only Slide 2 (2-case grid with Challenge/Solution/Impact)
  - `4` = Only Slide 1 (4-case overview with metrics)

---

## Response Format

### Success Response (200 OK)

**Content-Type:** `application/vnd.openxmlformats-officedocument.presentationml.presentation`

**Headers:**
```
Content-Disposition: attachment; filename="MedTech_Solutions_4-cases_20251111_120530.pptx"
Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation
```

**Body:** Binary PPTX file (streaming response)

**Filename Format:** `{company_name}_{suffix}_{timestamp}.pptx`
- `suffix`: `"all-slides"` (type=0) or `"N-cases"` (type=1/2/4)
- Spaces replaced with underscores
- Special characters removed
- Timestamp: `YYYYMMDD_HHMMSS`

**Examples:**
- `MedTech_Solutions_4-cases_20251111_120530.pptx` (4 cases, 1 slide)
- `MedTech_Solutions_1-cases_20251111_120530.pptx` (1 case, 1 slide)
- `MedTech_Solutions_all-slides_20251111_120530.pptx` (all slides)

---

## Error Responses

### 400 Bad Request - Invalid Input

**When:** Invalid `presentation_type` value

```json
{
  "error": "Invalid request",
  "detail": "presentation_type must be 0, 1, 2, or 4"
}
```

### 401 Unauthorized - Missing API Key

**When:** OpenAI API key not configured on server

```json
{
  "error": "OpenAI API key not configured"
}
```

### 422 Unprocessable Entity - Validation Error

**When:** Request body doesn't match schema (automatic Pydantic validation)

```json
{
  "detail": [
    {
      "type": "string_too_short",
      "loc": ["body", "company_description"],
      "msg": "String should have at least 10 characters",
      "input": "Short"
    }
  ]
}
```

**Common validation errors:**
- `company_name` is empty or missing
- `company_description` is too short (<10 chars) or too long (>2000 chars)
- `presentation_type` is not 0, 1, 2, or 4
- Missing required fields

### 500 Internal Server Error - Generation Failed

**When:** Unexpected error during generation

```json
{
  "error": "Presentation generation failed",
  "detail": "Data file not found: /app/data/case_studies_complete.json"
}
```

### 502 Bad Gateway - External Service Failure

**When:** OpenAI API fails or times out

```json
{
  "error": "External service error",
  "detail": "OpenAI API request failed"
}
```

---

## Frontend Implementation Examples

### React + TypeScript + Fetch

```typescript
interface GenerateRequest {
  company_name: string;
  company_description: string;
  presentation_type: 0 | 1 | 2 | 4;
}

async function generatePresentation(
  companyName: string,
  companyDescription: string,
  presentationType: 0 | 1 | 2 | 4
): Promise<void> {
  const API_URL = 'https://your-app.up.railway.app';

  try {
    const response = await fetch(`${API_URL}/api/generate`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        company_name: companyName,
        company_description: companyDescription,
        presentation_type: presentationType,
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.detail || error.error || 'Generation failed');
    }

    // Extract filename from Content-Disposition header
    const contentDisposition = response.headers.get('Content-Disposition');
    const filename = contentDisposition
      ?.match(/filename="(.+)"/)?.[1]
      || `presentation_${Date.now()}.pptx`;

    // Download the file
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);

    console.log('✅ Presentation downloaded successfully');
  } catch (error) {
    console.error('❌ Generation failed:', error);
    throw error;
  }
}

// Usage
await generatePresentation(
  'MedTech Solutions',
  'Healthcare SaaS company focused on patient engagement.',
  4
);
```

### React Component Example

```tsx
import { useState } from 'react';

export function PresentationGenerator() {
  const [companyName, setCompanyName] = useState('');
  const [description, setDescription] = useState('');
  const [presentationType, setPresentationType] = useState<0 | 1 | 2 | 4>(4);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleGenerate = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError(null);

    try {
      const response = await fetch('https://your-app.up.railway.app/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          company_name: companyName,
          company_description: description,
          presentation_type: presentationType,
        }),
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.detail || err.error || 'Failed to generate');
      }

      // Download file
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${companyName.replace(/\s+/g, '_')}_presentation.pptx`;
      a.click();
      window.URL.revokeObjectURL(url);

      alert('✅ Presentation downloaded!');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <form onSubmit={handleGenerate}>
      <input
        type="text"
        placeholder="Company Name"
        value={companyName}
        onChange={(e) => setCompanyName(e.target.value)}
        required
        minLength={1}
        maxLength={200}
      />

      <textarea
        placeholder="Company Description (min 10 characters)"
        value={description}
        onChange={(e) => setDescription(e.target.value)}
        required
        minLength={10}
        maxLength={2000}
        rows={4}
      />

      <select
        value={presentationType}
        onChange={(e) => setPresentationType(Number(e.target.value) as 0 | 1 | 2 | 4)}
      >
        <option value={0}>All Slides</option>
        <option value={1}>1 Case Study</option>
        <option value={2}>2 Case Studies</option>
        <option value={4}>4 Case Studies</option>
      </select>

      <button type="submit" disabled={loading}>
        {loading ? 'Generating...' : 'Generate Presentation'}
      </button>

      {error && <div className="error">{error}</div>}
    </form>
  );
}
```

### Next.js API Route (Server-Side Proxy)

```typescript
// app/api/generate-presentation/route.ts
import { NextResponse } from 'next/server';

export async function POST(request: Request) {
  try {
    const body = await request.json();

    // Forward request to backend API
    const response = await fetch('https://your-app.up.railway.app/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const error = await response.json();
      return NextResponse.json(error, { status: response.status });
    }

    // Stream the binary response back to client
    const blob = await response.blob();
    return new NextResponse(blob, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': response.headers.get('Content-Disposition') || '',
      },
    });
  } catch (error) {
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    );
  }
}
```

### Axios Example

```typescript
import axios from 'axios';

async function downloadPresentation(
  companyName: string,
  description: string,
  type: 0 | 1 | 2 | 4
) {
  try {
    const response = await axios.post(
      'https://your-app.up.railway.app/api/generate',
      {
        company_name: companyName,
        company_description: description,
        presentation_type: type,
      },
      {
        responseType: 'blob', // Important for binary data
      }
    );

    // Extract filename
    const contentDisposition = response.headers['content-disposition'];
    const filename = contentDisposition?.match(/filename="(.+)"/)?.[1]
      || 'presentation.pptx';

    // Download
    const url = window.URL.createObjectURL(new Blob([response.data]));
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();
    link.remove();
  } catch (error) {
    if (axios.isAxiosError(error)) {
      console.error('API Error:', error.response?.data);
      throw new Error(error.response?.data?.detail || 'Generation failed');
    }
    throw error;
  }
}
```

---

## Performance Expectations

| Metric | Expected Value |
|--------|----------------|
| **Response Time** | 3-8 seconds |
| **AI Selection** | 2-5 seconds |
| **PPTX Generation** | 1-3 seconds |
| **File Size** | 2-4 MB |
| **Timeout Recommendation** | 30 seconds |

**Notes:**
- Response is **streamed** (no buffering delay)
- **No polling required** - single request/response
- Add loading UI for 3-8 second wait time
- Consider timeout of 30s to handle slow AI responses

---

## Error Handling Best Practices

### Recommended User Messages

```typescript
function getUserFriendlyError(status: number, detail: string): string {
  switch (status) {
    case 400:
      return 'Invalid input. Please check your form values.';
    case 401:
      return 'Service temporarily unavailable. Please try again later.';
    case 422:
      if (detail.includes('company_description')) {
        return 'Company description must be at least 10 characters long.';
      }
      if (detail.includes('presentation_type')) {
        return 'Please select a valid presentation type (1, 2, or 4 cases).';
      }
      return 'Please check your input and try again.';
    case 500:
      return 'Server error. Our team has been notified. Please try again in a few minutes.';
    case 502:
      return 'External service error. Please try again.';
    default:
      return 'An unexpected error occurred. Please try again.';
  }
}
```

### Retry Logic

```typescript
async function generateWithRetry(
  data: GenerateRequest,
  maxRetries = 2
): Promise<Blob> {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await fetch('https://your-app.up.railway.app/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });

      if (!response.ok) {
        const error = await response.json();
        // Don't retry validation errors (422) or bad requests (400)
        if (response.status === 422 || response.status === 400) {
          throw new Error(error.detail || error.error);
        }
        throw new Error(`HTTP ${response.status}`);
      }

      return await response.blob();
    } catch (error) {
      if (attempt === maxRetries) throw error;

      // Exponential backoff: 1s, 2s, 4s...
      const delay = 1000 * Math.pow(2, attempt - 1);
      await new Promise(resolve => setTimeout(resolve, delay));
    }
  }
  throw new Error('Max retries exceeded');
}
```

---

## Testing the API

### Health Check Endpoint

```bash
GET /health

Response:
{
  "status": "healthy",
  "service": "case-study-generator"
}
```

Use this for uptime monitoring and deployment verification.

### Interactive API Documentation

Visit `https://your-app.up.railway.app/docs` for:
- **Swagger UI** - Interactive API testing
- **Try it out** - Test endpoint directly in browser
- **Schema validation** - See request/response formats

---

## CORS Configuration

**Current:** Allows all origins (`*`)

**Production Recommendation:** Update `api/app.py` to whitelist your frontend domain:

```python
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://your-frontend-domain.com"],
    allow_credentials=True,
    allow_methods=["POST"],
    allow_headers=["Content-Type"],
)
```

---

## Common Issues & Solutions

### Issue: Request hangs or times out
**Solution:** Increase fetch timeout to 30 seconds. The AI selection takes 2-5 seconds.

### Issue: CORS error in browser console
**Solution:** Backend CORS is configured. Check if you're sending correct headers.

### Issue: 422 validation error on valid input
**Solution:** Check that `presentation_type` is **number** (0, 1, 2, 4), not string ("0", "1", "2", "4")

### Issue: Downloaded file is corrupt
**Solution:** Ensure `responseType: 'blob'` in fetch/axios. Don't try to parse binary as JSON.

### Issue: Filename is generic "presentation.pptx"
**Solution:** Extract from `Content-Disposition` header (see examples above)

---

## Support

**API Issues:** Check Railway deployment logs
**Frontend Integration:** See examples above
**Questions:** Review `/docs` endpoint for interactive testing

**API Version:** 1.0.0
**Last Updated:** 2025-11-11
