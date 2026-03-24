# TAS Weekly Consolidated Deck Generator

Generate the TAS Weekly Consolidated HTML report with 6 tabs covering TAS + Aliado.

**Output**: Single .html file, Chart.js charts, responsive tables, mobile-friendly.
**Trigger**: weekly, weekly review, weekly deck, weekly W{N}
**Save to**: ~/Downloads/TAS_Weekly_Consolidated_W{N}_{Year}.html
**Reference file**: ~/Downloads/TAS_Weekly_Consolidated_W12_2026.html (always open and replicate)

---

## Tabs
1. **Vista Ejecutiva** — TAS+Aliado consolidado (21 columnas)
2. **Funnel TAS** — Por grupo, solo TAS
3. **Funnel Aliado** — Por grupo, solo Aliado
4. **Charts** — 7 gráficas en grid 2x2 (25% cada una del frame)
5. **Scorecard** — Tabla interactiva por grupo, filtros General/TAS/Aliado + pills de métricas
6. *(Pendiente)* Diagnóstico / Plan de acción

---

## 21-Column Table Structure

### Sub-header row (2-level)
| Grupo header | ← 5 Semanas → | ← Mensual → | MTD x2 | Comp | Best | RR x2 | Target x2 (purple) |
|---|---|---|---|---|---|---|---|

### Column labels
`Métrica | w-5 | w-4 | w-3 | w-2 | w-1 | WoW | Oct | Nov | Dic | Ene | Feb | Mar* | Feb MTD | Mar MTD | MoM | Best | RR 7d | RR Seas | Target BP | % BP`

### Weekly periods (Mon–Sat, excluding Domingos y feriados)
- w-1 = Mar 16–21 (Lunes Mar 16 fue feriado → 5 días hábiles)
- w-2 = Mar 9–14
- w-3 = Mar 2–8
- w-4 = Feb 23–28
- w-5 = Feb 16–22
- **Días hábiles Marzo 2026 = 25** (Lunes Mar 16 = feriado)

### Calendar labels above weekly columns
Show the date range (e.g. "16–21 Mar") above w-1, w-2 etc. in a sub-row.

---

## Row Structure per Group (8 rows)

| # | Métrica | CSS class | Notes |
|---|---|---|---|
| 1 | Quotes | dr | |
| 2 | Made | dr | |
| 3 | Approved | dr | |
| 4 | Purchased | dr bl | Bold, blue highlight |
| 5 | Q→M% | dr | |
| 6 | A→P% (purchase) | dr | Current: Purchased / Approved |
| 7 | A→P% cohort cerrada | dr pu | Purple. Of approvals ≥7 days old, % with purchase within 7d of approval |
| 8 | Q→P% | dr bl qp | Bold blue |

### Group separator rows (tr.grp)
- Span 21 cols. Show grupo name + MoM deltas: `Grupo → Q +X% | M +X% | P +X%`
- **MoM = Feb MTD (día 20) vs Mar MTD (día 20)** — NEVER full Feb vs partial Mar
- TAS Consolidado bg: `#004E98`. Aliado Consolidado bg: `#3B7DDD`.
- Each grupo header color: `#1B2A4A`

### Vista Ejecutiva extra rows (after groups)
- `Compras / FTE TAS` — Purchased / FTE count (White Label from People Cost CSV)
- Separator `tr.sp`
- `Quotes / día` — Quotes / business days
- `Mades / día` — Made / business days
- `Purchases / día` — Purchased / business days
- `Células TAS` — active TAS cells (cells with ≥1 quote)
- `FTEs TAS` — from People Cost CSV, Pilar = "White Label"

---

## WoW Calculation Rule (CRITICAL)
**WoW MUST be computed from the values already in w-1 and w-2 columns — NEVER from raw Python date queries.**
- If w-1 > w-2 → WoW is POSITIVE
- Volume metrics: `(w1 - w2) / w2 * 100` as %
- Rate metrics (Q→M%, A→P%, Q→P%): `w1 - w2` as pp (percentage points)
- Use a JS/Python function that reads the rendered values, not a fresh query

### Monthly MoM Rule
- MoM in monthly columns: show % growth vs previous month below the value
- Oct = base (no delta). Nov shows % vs Oct, Dic vs Nov, etc.
- Green if positive, red if negative

---

## MoM Header Label Format
Each grupo header shows: `Grupo → Q +X% | M +X% | P +X%`
where X% = (Mar MTD day 20 – Feb MTD day 20) / Feb MTD day 20

---

## Inspector / Capacity Logic
- Default: **1 inspector per active TAS cell**
- Default: **1 inspector per active ALIADO cell**
- Exceptions:
  - ISMO AGUASCALIENTES (TAS) + ALIADO AGUASCALIENTES = **2 inspectors total**
  - ISMO LEON (TAS) + ALIADO LEON = **2 inspectors total**
- Capacity = inspectors × 10 mades/day × business days in period
- Exclude TEST-TEST

---

## Compras / FTE
- FTE source: `~/Downloads/People Cost - Liquidity Marketplace - Sheet1.csv`
- Filter: column `AE = "White Label"` (Pilar column)
- Compras/FTE = Total TAS Purchases / FTE count
- Show in Vista Ejecutiva bottom rows

---

## Run Rate Calculations
- **RR 7d** = (last 7 days purchases / 7) × business days remaining in month
- **RR Seasonality** = MTD purchases + (weighted average of last 5 weeks' daily pattern × remaining days)
- Formula displayed in footer: `RR 7d = (ult 7 días / 7) × 31 | RR Season = MTD + (prom ponderado 5 semanas × días restantes)`

---

## KPI Definitions (Vista Ejecutiva footer rows)
- `Avg P / Unit TAS` — complete in RR 7d and RR Seas columns (run rate for price/unit)
- `Best` column — for daily rates, show the best individual month value (not weekly)

---

## Target BP
- Source: `~/Downloads/Liquidity BP 2026 - Liquidity.csv`
- Column: Target BP per month
- `% BP` = Actual / Target × 100
- Purple color scheme for these 2 columns

---

## Design System
```
--dark:   #1A1A2E   (dark bg)
--hdr:    #2B478B   (table header)
--hdr2:   #3A569A   (sub-header)
--g:      #22C55E   (green / positive)
--r:      #EF4444   (red / negative)
--a:      #F59E0B   (amber / neutral)
--p:      #8B5CF6   (purple / cohort rows)
--lg:     #F8FAFC   (light gray row alt)
--brd:    #E2E8F0   (borders)
--sep:    #E8ECF1   (separator lines)
--grupo:  #1B2A4A   (grupo header bg)
--ali:    #3B7DDD   (aliado color)
--target: #7C3AED   (target purple)
```
Font: system-ui 11px. Sticky header/tabs/first-col.
Zero values: color #CBD5E1 (light gray, not bold).

---

## Charts Tab (7 charts, 25% frame each in 2×2 grid + 3rd row)

| # | Chart | Type | Notes |
|---|---|---|---|
| 1 | Compras Mensuales Oct→Mar | Stacked bar (TAS+Aliado) | RR dashed line, Feb ref line, labels on bars, growth annotation Oct→Mar RR |
| 2 | Waterfall Total MoM | Bar (Feb MTD → ΔTAS → ΔAliado → Mar MTD) | Green/red deltas |
| 3 | Waterfall TAS Funnel | Grouped bar Q/M/A/P Feb vs Mar MTD | Delta labels, active cells annotation below: "23 cél.→26 (+3) \| Q/cél: 67→74 (+10%)" |
| 4 | Waterfall Aliado Funnel | Same structure | Annotation: "8 aliados→15 (+7) \| Q/aliado: 209→109 (-48%)" |
| 5 | Conversiones TAS | 4 lines Q→M/M→A/A→P/Q→P monthly | Separate from Aliado |
| 6 | Conversiones Aliado | Same 4 lines | Separate from TAS |
| 7 | Aporte por Región | Horizontal stacked bar (last item) | Lerma/GDL/QRO/MTY/CDMX |

### Region Mapping
| Región | Grupos / Células |
|---|---|
| Kavak GDL | GRUPO PREMIER (Culiacán/Mazatlán/Hermosillo), GRUPO PLASENCIA |
| Kavak QRO | ISMO LEON, ALIADO LEON, ISMO AGUASCALIENTES, ALIADO AGUASCALIENTES, GRUPO POTOSINA |
| Kavak CDMX | SONI PACHUCA (+ anything explicitly mapped) |
| Kavak MTY | GRUPO WECARS, ALIADO WECARS, GRUPO MISOL, ALIADO MISOL, GRUPO AUTOPOLIS |
| Kavak Lerma | Everything else |

---

## Scorecard Tab (new design, ref image)

### Layout
```
Grupos
{N} grupos · {Q} quotes · {P} compras · Mar 2026

[General] [TAS] [Aliado]   ← filter pills

[Purchased] [Quotes] [Made] [Approved] [Q→M] [M→A] [A→P] [Q→P]  ← metric pills
```

### Table columns
`ENTIDAD | RUN RATE (blue bg #DBEAFE) | MAR 26* | FEB 26 | ENE 26 | DIC 25 | NOV 25 | OCT 25`

- Rows sorted by RUN RATE descending (re-sort on metric switch)
- Zero values: #CBD5E1
- RR = Mar* / 16 × 25 (16 elapsed business days, 25 total in March)
- Filter tabs update visible rows AND subtitle counts
- Metric pills switch all cell values + re-sort

---

## Data Sourcing
- **Pre-Mar 2026**: CSV `~/Downloads/Summary MKP _ TAS & BULK - Snapshot_Diario (7).csv`
- **Mar 2026+**: Excel `~/Downloads/TAS_Continental_ALL.xlsx`
- **Purchases**: filter `purchased == 1`, group by `purchase_date` (NEVER by `Fecha de creación`)
- **Quotes/Made/Approved**: group by `Fecha de creación`
- **Channel**: column `BULK\nTAS\nALIADO` → values "TAS", "ALIADO"
- **Exclude**: `Célula == "TEST-TEST"`

---

## Known Cell Mappings (Channel)
| Grupo | Channel | Notes |
|---|---|---|
| GRUPO CONTINENTAL | TAS | Main CDMX group, ~15 células |
| GRUPO PREMIER | TAS | Culiacán / Mazatlán / Hermosillo |
| GRUPO PLASENCIA | TAS | ~14 agencies GDL |
| GRUPO ANDRADE | TAS | |
| GRUPO ISMO TLALNEPANTLA | TAS | |
| GRUPO ISMO LEON | TAS | 2 inspectors (TAS+Aliado León) |
| GRUPO ISMO AGUASCALIENTES | TAS | 2 inspectors (TAS+Aliado AGS) |
| ALIADO LEON | ALIADO | Omoda brand |
| ALIADO COAPA | ALIADO | |
| GRUPO POTOSINA | ALIADO | Keep in scorecard |
| GRUPO WECARS | TAS | → Kavak MTY region |
| GRUPO MISOL | TAS | → Kavak MTY region |
| GRUPO AUTOPOLIS | TAS | → Kavak MTY region |
| GRUPO SONI | TAS | includes Pachuca → Kavak CDMX |
