# Case Study Presentation Generator

## Quick Start (Windows PowerShell)

```powershell
cd "C:\Users\joaop\Documents\Augusta Labs\Cases Study Slide Optimization\case_study_generator"
conda deactivate
uv pip install --system -r requirements.txt

# Example: Healthcare tech company
uv run python -m src.cli `
  --company-name "MedTech Solutions" `
  --company-description "Healthcare SaaS company with 200 employees focused on patient engagement and clinical workflow automation. Key challenges: improving care coordination, reducing administrative burden, and scaling AI capabilities."

# Example: Logistics/Supply chain company  
uv run python -m src.cli `
  --company-name "LogiFlow Systems" `
  --company-description "B2B logistics platform with 500 employees specializing in freight management and route optimization. Seeking to reduce operational costs by 30% and implement real-time tracking with AI."

# Example: Financial services
uv run python -m src.cli `
  --company-name "FinServe Pro" `
  --company-description "Enterprise fintech serving mid-market banks with 150 employees. Focus on multiproduct sales automation, client intelligence, and regulatory compliance. Target: 40% faster proposal generation."
```

**What gets generated:**
- ✅ 4 AI-selected case studies with images
- ✅ Metrics (e.g., "320 Hours Saved", "-70% Time-to-Insights")
- ✅ Categories (Infrastructure, Healthcare, Software, etc.)
- ✅ Professional titles and descriptions
- ✅ Company logos/images for each case study

## About

Automates case study presentation generation. Give it a company description, it uses AI to select the 4 most relevant case studies from the database and generates a PowerPoint presentation.

## How It Works

1. **Parse**: Reads 46 case studies from Excel
2. **Select**: GPT-4o-mini picks the 4 most relevant
3. **Generate**: Fills PowerPoint template

## Setup

Install with uv (fast):
```powershell
uv pip install --system -r requirements.txt
```

Create `.env` file:
```
OPENAI_API_KEY=your_key_here
```

### Options

- `--company-name`: Required
- `--company-description`: Required  
- `--output-dir`: Default `output/`
- `--template`: Default `templates/Case studies Template (1).pptx`
- `--data`: Default `data/Case Study Matrix.xlsx`
- `--api-key`: Overrides environment variable

## Template Placeholders

Configure in `src/config.py`:

- `{{slide_case_studies_subtitle}}` - Slide subtitle
- `{{case_study_n_name}}` - Organization + Title (n=1-4)
- `{{case_study_n_image}}` - Image placeholder (n=1-4)
- `{{n1}}` to `{{n4}}` - Metrics
- `{{company_name}}`, `{{company_description}}`, `{{generation_date}}`

## Project Structure

```
case_study_generator/
├── src/
│   ├── excel_parser.py      # Extract data from Excel
│   ├── ai_selector.py        # AI selection
│   ├── pptx_generator.py     # PowerPoint generation
│   ├── config.py             # Template config
│   └── cli.py                # CLI
├── templates/                # PowerPoint templates
├── data/                     # Excel data
└── output/                   # Generated presentations
```

## Available Case Studies (46 total)

### By Sector

**Healthcare & Life Sciences**
- Astrazeneca (Intelligence for Specialists)
- Maia Clínica (Voice AI, Clinical Documentation)
- Sword Health (Shortlisting Engine)
- Vet AI (Diagnostic Automation)
- Medicare (Revenue Recovery)
- Pet24 (Scheduling)

**Financial Services**
- Banco BIG (Investment Recommender, Client Scoring, CRM)
- Lince Capital (Intelligence & Prospecting, Investor KYC)
- Sage (Compliance Coach, Cash-Flow Forecasting)
- Millennium BCP (Management Assistant)

**Logistics & Operations**
- Ferrovial (Win-Rate Optimization)
- AddVolt (Enterprise Knowledge Retrieval)
- Portir (Optimization Engine)

**Retail & Consumer**
- Rádio Popular (Margin Optimizer)
- Bene (Demand & Inventory Forecast)
- Farmácias Grupo STS (AutoML)

**Media & Entertainment**
- 24 horas (Distribution Engine)
- Media Capital (Sentiment Intelligence)
- SportTV (Renewal Engine)

**Technology & Software**
- Clever Advertising (Discovery Engine)
- DeskSkill (Autonomous Recruitment)
- Gist (Ops Control Platform)

**Public Sector**
- SPMS (Compliance Copilot NHS)
- Câmara de Matosinhos (Intelligence Dashboard)
- FAP (Chatbot Copilot)

**Other**
- Tripwix (Listing Velocity)
- Grupo Embrace (Velocity Automation)
- Quadrante (Mining & Exploration)

## AI Selection

The AI analyzes:
- Industry alignment
- Similar challenges
- Complementary perspectives
- Solution relevance

Returns 4 diverse, compelling examples.

## Technical Details

**Dependencies:**
- python-pptx 0.6.23 - PowerPoint manipulation
- openpyxl 3.1.2 - Excel reading
- openai >=1.99.5 - AI API
- click 8.1.7 - CLI framework
- python-dotenv >=1.1.0 - Environment variables

**Key Decisions:**
- Placeholder pattern: `{{variable}}` - visual and conflict-free
- GPT-4o-mini - cost-effective with excellent results
- Run-level text replacement - preserves formatting
- Centralized config - easy template customization

## Troubleshooting

**"OpenAI API key not found"**
- Create `.env` with `OPENAI_API_KEY=your_key`
- Or use `--api-key`

**"Data file not found"**
- Check `Case Study Matrix.xlsx` in `data/`
- Or specify with `--data`

**"Template file not found"**
- Check template in `templates/`
- Or specify with `--template`

**"Module not found"**
- Navigate to project directory first
- `cd "C:\Users\joaop\Documents\Augusta Labs\Cases Study Slide Optimization\case_study_generator"`

