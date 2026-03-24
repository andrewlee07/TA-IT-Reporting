# TeacherActive Exec Reporting App v1

Internal Next.js app for workbook-driven executive IT reporting. The app ingests a locked Excel template, validates it against the `v4` contract, stores the uploaded workbook and normalized snapshot, renders the 13-page exec pack, and exports report assets as PNG, PDF, or PowerPoint.

## Stack

- `Next.js 16` + `React 19` + `TypeScript`
- `Prisma 7` + PostgreSQL JSONB snapshots
- Local or S3 object storage for uploaded workbooks and generated exports
- `Playwright` for pixel-matched PNG/PDF/PPTX export generation
- `xlsx`, `jszip`, and XML inspection for workbook ingestion and contract validation

## Workbook Contract

The app only accepts the locked `v3` template.

- `Template Key = IT_EXEC_TEMPLATE_V4`
- `Template Version = 4`

### Required parsed sheets

- `Periods`
- `Entities`
- `Office_Locations`
- `INPUT_Office_Network_Avail`
- `INPUT_Service_Availability`
- `INPUT_Support_Operations`
- `INPUT_Top_Oldest_Tickets`
- `INPUT_Security_Patching`
- `INPUT_Assets_Lifecycle`
- `INPUT_Change_Release`
- `INPUT_Dev_Delivery`
- `INPUT_Project_Portfolio`
- `INPUT_Rolling_Roadmap`
- `INPUT_Gantt_Workstreams`
- `INPUT_Gantt_Milestones`
- `INPUT_Budget_Commercials`
- `INPUT_Top_Risks`
- `INPUT_Narrative_Notes`

### Office network additions in v2

- `Office_Locations` must contain table `TOfficeLocations`
- `INPUT_Office_Network_Avail` must contain table `TOfficeNetworkAvailability`
- `INPUT_Service_Availability` must not contain manual `Network` rows

Overall Network performance is derived from office rows per reporting month:

- `Availability %` = arithmetic mean of in-scope office availability
- `Outage Minutes` = sum of office outage minutes
- `Major Incidents` = sum of office major incidents
- Derived office metrics also include `perfectOffices`, `below99_9Offices`, `below99Offices`, and `worstOffice`

### Portfolio Gantt additions in v3

- `Periods` must include `Report Cut-Off Date`
- `INPUT_Gantt_Workstreams` must contain table `TPortfolioGanttWorkstreams`
- `INPUT_Gantt_Milestones` must contain table `TPortfolioGanttMilestones`
- `INPUT_Gantt_Workstreams.Domain` must be one of:
  - `Infrastructure`
  - `End-user computing`
  - `Security & compliance`
  - `Applications & data`
  - `Product / development`
  - `Business transformation`
- `Portfolio Gantt` renders a 12-week window starting from the first Monday on or after the selected reporting month
- The orange vertical line uses the stored `Periods.Report Cut-Off Date`

Note: the original long-form Gantt sheet names exceeded Excel's 31-character worksheet limit, so the workbook uses the Excel-safe sheet names `INPUT_Gantt_Workstreams` and `INPUT_Gantt_Milestones`.

## Getting Started

1. Install dependencies:

```bash
npm install
```

2. Copy env vars:

```bash
cp .env.example .env
```

3. Start PostgreSQL and set `DATABASE_URL`.

4. Generate the Prisma client and run migrations:

```bash
npm run db:generate
npm run db:migrate
```

5. Start the app:

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

## Template Files

- Bundled v4 fixture: [fixtures/IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx](/Users/andrewlee/GitHub/TA-IT Reporting Claude/fixtures/IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx)
- Downloadable template: [public/templates/IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx](/Users/andrewlee/GitHub/TA-IT Reporting Claude/public/templates/IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx)

To regenerate the workbook after changing the upgrade script:

```bash
npm run template:upgrade
```

To generate a JSON snapshot from the bundled workbook:

```bash
npm run snapshot:demo
```

## API

- `POST /api/reports`
  - Upload a workbook, validate it, parse it, persist the workbook, and store the normalized snapshot
- `GET /api/reports/:id`
  - Load report metadata and its stored snapshot
- `POST /api/reports/:id/exports`
  - Generate:
    - `page-png`
    - `block-png`
    - `full-pdf`
    - `full-pptx`

## Verification

Typecheck:

```bash
npm run typecheck
```

Lint:

```bash
npm run lint
```

Unit tests:

```bash
npm test
```

E2E smoke tests:

```bash
npm run test:e2e
```
