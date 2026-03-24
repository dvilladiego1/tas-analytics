# Project: TAS Analytics & Report Generator

## User: Daniel Villadiego

## Key Skillset: TAS Monthly Business Review (MBR)

### Quick Reference
- Full skillset doc: `/Users/danielvilladiego/Downloads/SKILLSET_TAS_MBR_Report.md`
- Definitive deck template: `/Users/danielvilladiego/Downloads/Kia Coapa Review Feb 2026.pptx`

### Data Sources
- **CSV Snapshot**: `Summary MKP _ TAS & BULK - Snapshot_Diario (N).csv` in Downloads (N increments; latest = 8)
- **Excel per cell**: `TAS_Continental_-_{Celula}_2026-{MM}.xlsx` in Downloads
- **Excel all cells**: `TAS_Continental_ALL.xlsx` in Downloads
- **Retail Supply CSV**: `MX KPIS [Oficial] - DB Supply (1).csv` in Downloads
  - Multi-block horizontal structure (5 blocks side-by-side, each with own date+region columns)
  - Column indices: registros_exitosos=3, schedule_confirmed=18, inspection_made=27, inspection_approved=38, net_purchases_retail=47
  - Filter `region="All"` for national totals. Daily aggregated (1 row per date+region).

### CSV Column Mapping
- Cell filter: column `Célula` (e.g., "CONTINENTAL COAPA", "ALIADO COAPA", "ALIADO TLAHUAC")
- Group filter: column `Grupo` (e.g., "GRUPO PLASENCIA", "GRUPO CONTINENTAL")
- Channel type: column `BULK\nTAS\nALIADO` → values: "TAS", "ALIADO", "BULK"
- Creation date: `Fecha de creación` (dd/mm/yyyy format) — para Quotes, Made, Approved
- Purchase date: `purchase_date` (dd/mm/yyyy format) — ⚠️ USAR PARA COMPRAS NETAS
- Funnel flags: `scheduled.1` (quote), `made.1` (inspected), `approved`, `purchased`

### ⚠️ Metodología de Conteo de Compras
**REGLA CRÍTICA**: "Compras netas del mes X" = registros con `purchase_date` en mes X.
- NO usar `Fecha de creación` para contar compras por mes.
- Un quote creado en Enero con purchase_date en Febrero cuenta como compra de FEBRERO.
- Aplica a: hero numbers, evolución histórica, compras diarias/semanales, A→P%, Q→P%.
- Quotes, Made, Approved sí se agrupan por `Fecha de creación`.

### Known Cell Mappings
| Cell Name | TAS Célula / Grupo | Aliado Célula(s) | Notas |
|-----------|-------------------|-----------------|-------|
| Kia Coapa | CONTINENTAL COAPA | ALIADO COAPA | TAS + Aliado |
| Kia Tláhuac | (same cell as Coapa) | ALIADO TLAHUAC | TAS + Aliado |
| Interlomas | CONTINENTAL INTERLOMAS | — | 100% TAS |
| Santa Fe | CONTINENTAL SANTA FE | — | — |
| Metepec | CONTINENTAL METEPEC | — | — |
| Patriotismo | CONTINENTAL PATRIOTISMO | — | — |
| Plasencia | Grupo = 'GRUPO PLASENCIA' | — | 100% TAS, ~14 agencias GDL |
| Premier | Grupo = 'GRUPO PREMIER' | — | 100% TAS, 3 células (Culiacán/Mazatlán/Hermosillo), Sinaloa/Sonora |
| Ismo | Grupo LIKE 'GRUPO ISMO%' | ALIADO LEON | TAS 3 células (Tlalnepantla/León/AGS) + Aliado León (OMODA) |
| Andrade | GRUPO ANDRADE | — | 100% TAS, 3 células (Aeropuerto/Azcapotzalco/Cuautitlán), CDMX |
| Soni | GRUPO SONI | — | 100% TAS, 2 células (Pachuca/Querétaro) |
| Potosina | GRUPO POTOSINA | — | 100% TAS, SLP. Override channel to TAS always |
| Wecars | GRUPO WECARS | — | MTY area |
| Autopolis | GRUPO AUTOPOLIS | — | MTY area |
| GP Auto | GRUPO GP AUTO | — | 1 célula |
| Misol | GRUPO MISOL | — | MTY area |
| Mega | — | ALIADO MEGA | Aliado only |
| Torres Corzo | — | ALIADO TORRES CORZO | Aliado, SLP area |
| Tollocan | — | ALIADO TOLLOCAN | Aliado, EdoMex |

### PPTX Design System (from definitive deck)
- **Slide size**: 13.333" x 7.500"
- **Dark bg**: #1A1A2E | **Brand blue**: #004E98 | **Table header**: #2B478B
- **Green**: #22C55E | **Red**: #EF4444 | **Amber**: #F59E0B | **Purple**: #8B5CF6
- **Chart series**: TAS=#1B2A4A, Aliado1=#3B7DDD, Aliado2=#00B48A
- **Insight boxes**: Highlights=#DCFCE7, Lowlights=#FEE2E2, Summary=#DBEAFE, Lectura=#E3EEFB
- **Fonts**: Title=24-36pt Bold, KPI=28-48pt Bold, Table=13pt, Insight=9-11pt, Footer=8-9pt

### Slide Structure (12 slides)
1. Title (dark bg)
2. Hero numbers — compras, meses crecimiento, meta
3. Executive Summary — 4 KPI cards + highlights/lowlights
4. Evolución Histórica — stacked bar chart + tables
5. Comparativo TAS (Ene vs Feb table)
6. Comparativo Aliado 1 (if applicable)
7. Comparativo Aliado 2 (if applicable)
8. Transition (optional)
9. Estacionalidad — weekly + daily tables
10. Diagnóstico Funnel — cascade waterfall
11. Plan de Acción
12. Closing / KPIs to Monitor

### Generation Prompt
When Daniel asks to generate an MBR for a new cell, use:
1. Read the CSV Snapshot
2. Filter by Célula name and channel type
3. Aggregate monthly funnel (quotes→made→approved→purchased)
4. Compute conversion rates and MoM deltas
5. Generate all slides following the design system above
6. Save to Downloads as `{Cell}_MBR_{Month}_{Year}.pptx`

### Scripts
- `gen_weekly_w12.py` — Weekly consolidated HTML report (5 tabs). Core data functions: `load_csv()`, `aggregate()`, `rr_7d()`.
- `gen_plan_mejora_abril.py` — Plan de Mejora Abril 2026 PPTX (8 slides, Retail vs TAS vs Aliado comparison + impact matrices)
- `build_mbr.py` — Monthly Business Review PPTX generator (per-cell). PPTX helper functions: `add_kpi_card()`, `add_insight_box()`, etc.
- `gen_premier_review.py` — Grupo Premier review generator
- `generate_sop_deck.py` — S&OP resource justification deck

### Dependencies
```
pip install python-pptx pandas openpyxl
```

---

## Data Conventions

### Date Fields
- **Purchases (compras netas)**: Always group by `purchase_date`, never by `Fecha de creación`.
- **Quotes / Made / Approved**: Group by `Fecha de creación`.
- **Date format**: dd/mm/yyyy in both CSV and Excel sources.

### Purchase Counting Rule
A purchase belongs to the month of its `purchase_date`, regardless of when the quote was created. A quote created in January with `purchase_date` in February is a February purchase.

### Funnel Flag Logic
- Quote = `scheduled.1` is truthy
- Inspected = `made.1` is truthy
- Approved = `approved` is truthy
- Purchased = `purchased` is truthy

### Purchase Filtering
When filtering purchases, always use the `purchased` flag field — never use `purchase_date` for filtering as it is mostly unpopulated and produces incorrect results. Use `purchase_date` only for grouping already-filtered purchases by month.

### Data Source Selection
- **TAS funnel data**: Use latest CSV Snapshot (`Snapshot_Diario (N).csv`). CSV8 covers through Mar 22, 2026.
- **Per-cell Excel**: Use when CSV Snapshot doesn't have the latest data or for cell-specific deep dives.
- **Retail benchmarks**: Use `MX KPIS [Oficial] - DB Supply (1).csv` — has data back to Jan 2025.

### Channel Overrides (applied in code)
- GRUPO POTOSINA: Always force channel = 'TAS' (CSV sometimes says ALIADO)
- ANDRADE AEROPUERTO: Force channel = 'TAS' when CSV says ALIADO
- SONI PACHUCA: Force channel = 'TAS' when CSV says ALIADO

---

## KPI Definitions

| Role / Stage | Primary KPI | Formula |
|---|---|---|
| Quotes (Q) | Quote volume | Count where `scheduled.1` is truthy, grouped by `Fecha de creación` |
| Made (M) | Inspection rate (Q→M%) | Made / Quotes × 100 |
| Approved (A) | Approval rate (M→A%) | Approved / Made × 100 |
| Purchased (P) | Purchase conversion (A→P%) | Purchased / Approved × 100 |
| End-to-end | Quote-to-Purchase (Q→P%) | Purchased / Quotes × 100 |
| Supply Agents | Purchase conversion (A→P%) | Purchased / Approved × 100 — **not** M→A% |
| Growth | MoM delta | (Current month − Previous month) / Previous month × 100 |
| Target | Meta achievement (%) | Actual purchases / Target × 100 |
| Channel mix | TAS vs Aliado share | Channel purchases / Total purchases × 100 |

**Important**: Always confirm KPI definitions with the user before generating evaluation or review documents. Supply Agents are measured by A→P%, not M→A%.

### Retail vs TAS Funnel Equivalence
| TAS Stage | Retail Equivalent | Comparable Rate |
|---|---|---|
| Quote (Q) | Schedule Confirmed | Retail S→M% ≈ TAS Q→M% |
| Made (M) | Inspection Made | Same |
| Approved (A) | Inspection Approved | Same |
| Purchased (P) | Net Purchases Retail | Same |

When comparing funnels, use Retail **S→M%** (Schedule-to-Made) as the equivalent of TAS **Q→M%** (Quote-to-Made). Both measure how many scheduled appointments become completed inspections.

### MTD Comparison Rule
When comparing Mar MTD vs Feb MTD, use same-period (same number of business days) to make volumes comparable. E.g., Mar 1-21 (~17 biz days) vs Feb 1-20 (~17 biz days). Conversion rates can use full-month data.

---

## Output Preferences

- **Comparison default**: Always show month-over-month (MoM) comparisons — never accumulated/cumulative — unless explicitly asked otherwise. Use green/red color coding for improvement/decline.
- **Concise communication**: Emails and summaries should be brief, data-driven, and action-oriented. 3-5 bullet points max unless told otherwise. No filler text.
- **File paths**: Always state the full output file path when saving any artifact (e.g., `/Users/danielvilladiego/Downloads/Kia_Coapa_MBR_Mar_2026.pptx`).
- **Language**: Reports and decks in Spanish. Code comments and CLAUDE.md in English.
- **Numbers**: Use comma as thousands separator for Spanish-facing outputs (e.g., 1,234). Percentages to one decimal (e.g., 45.2%).
- **Charts**: Always include data labels on bars/columns. Use the brand color palette from the design system.

---

## Document Generation

- **Always match reference styling**: When generating HTML decks or presentations, always match styling from the most recent reference file the user provides. Ask for a reference file if none is given. Never assume default styling.
- **Output file path**: Always state the full output file path clearly at the end of generation (e.g., `/Users/danielvilladiego/Downloads/Kia_Coapa_MBR_Mar_2026.pptx`).

---

## Styling

- **Always match reference files**: Every generated PPTX must follow the design system defined in the PPTX Design System section above. When in doubt, open the definitive deck (`Kia Coapa Review Feb 2026.pptx`) and replicate its layout.
- **Slide dimensions**: 13.333" × 7.500" — never deviate.
- **Color palette**: Use only the defined hex values for backgrounds, text, tables, and charts. Do not introduce new colors.
- **Font sizes**: Follow the ranges in the design system (Title 24-36pt, KPI 28-48pt, Table 13pt, Insight 9-11pt, Footer 8-9pt).
- **Insight boxes**: Use the defined background colors — green for highlights, red for lowlights, blue for summary/lectura.
- **Table headers**: Always use #2B478B background with white text.
- **Consistency**: If a new cell MBR is generated, it must be visually indistinguishable in style from existing decks.
