# Data Cleaning Pipeline

## VAT Zero-Rating, Nutrition-Relevant Food Price Inflation, and Healthy Diet Affordability in South Africa (2008–2025)

**Student:** Rochelle van der Walt (49218786)  
**Supervisor:** Prof. Waldo Krugell  
**Institution:** North-West University (NWU)

---

## Pipeline Overview

This repository contains the complete data cleaning pipeline that transforms raw Stats SA CPI data, FAO food price indices, and PMBEJD affordability data into the audited master dataset (`Food_VAT_EViews_Master_2008_2025_v3-3.xlsx`) used for econometric analysis.

### Pipeline Execution Order

```
Step 1: build_master_data.py
  ├── Reads raw CPI indices (COICOP 8-digit, 2008-2025)
  ├── Reads CPI average prices (Rand, 2017-2025)
  ├── Reads kJ energy density database
  ├── Reads FAO food price indices
  ├── Reads PMBEJD affordability data
  └── Outputs: intermediate .pkl and .json files

Step 2: build_excel.py
  ├── Same source data as Step 1
  └── Outputs: Food_VAT_EViews_Master_2008_2025.xlsx (v1)

Step 3: build_final_excel.py
  ├── Reads: Food_VAT_EViews_Master_2008_2025.xlsx + classification_data.json
  └── Outputs: Food_VAT_EViews_Master_2008_2025_v2.xlsx

Step 4: fix_categories.py + fix_eviews_master.py
  ├── Reads: v2 Excel file + classification_data.json
  └── Outputs: Food_VAT_EViews_Master_2008_2025_v3.xlsx

Step 5: build_phd_dataset.py
  ├── Reads: Food_VAT_EViews_Master_2008_2025_v3-3.xlsx (with manual corrections)
  └── Outputs: PhD_EViews_Master_2008_2025.xlsx (final audited, 14 sheets)
```

## Repository Structure

```
├── scripts/                    # Python data cleaning scripts
│   ├── build_master_data.py    # Step 1: Build master data from raw sources
│   ├── build_excel.py          # Step 2: Create initial Excel workbook
│   ├── build_final_excel.py    # Step 3: Add classifications and derived sheets
│   ├── fix_categories.py       # Step 4a: Fix food category classifications
│   ├── fix_eviews_master.py    # Step 4b: Fix EViews master sheet
│   └── build_phd_dataset.py    # Step 5: Build final audited PhD dataset
│
├── source_data/                # Raw input data files
│   ├── EXCEL-CPI-COICOP-2018-8digit-202512.xlsx   # Stats SA CPI indices
│   ├── CPI_Average-Prices_All-urban-202512.xlsx    # Stats SA average prices
│   ├── food_price_indices_data_feb-2.xlsx          # FAO food price indices
│   ├── pmbejd_affordability_2025-2.xlsx            # PMBEJD affordability data
│   ├── pmbejd_affordability_extended.xlsx          # PMBEJD extended series
│   ├── kj_per_100g_database.json                   # Energy density (kJ/100g)
│   └── classification_data.json                    # Food classification data
│
└── output_data/                # Output dataset
    └── Food_VAT_EViews_Master_2008_2025_v3-3.xlsx  # Final master dataset
```

## Source Data Description

| File | Source | Description |
|------|--------|-------------|
| `EXCEL-CPI-COICOP-2018-8digit-202512.xlsx` | Stats SA P0141 | 8-digit COICOP CPI indices, base Dec 2024=100, monthly 2008–2025 |
| `CPI_Average-Prices_All-urban-202512.xlsx` | Stats SA P0141 | Average consumer prices in Rands, all urban areas, 2017–2025 |
| `food_price_indices_data_feb-2.xlsx` | FAO | FAO Food Price Index — nominal monthly indices by food group |
| `pmbejd_affordability_2025-2.xlsx` | PMBEJD | Pietermaritzburg Economic Justice & Dignity food basket costs |
| `pmbejd_affordability_extended.xlsx` | PMBEJD | Extended PMBEJD affordability time series |
| `kj_per_100g_database.json` | SAFOODS | Energy density values (kJ per 100g) from SA Food Composition Database |
| `classification_data.json` | Derived | VAT zero-rated vs standard-rated food classification lookup |

## Output Dataset Sheets

The `Food_VAT_EViews_Master_2008_2025_v3-3.xlsx` contains 14 sheets:

1. **Product_Lookup** — 134 food products with COICOP codes, VAT status, energy density
2. **CPI_Indices** — Monthly CPI indices for 134 food items (2008–2025)
3. **Price_per_100kJ** — Price per 100 kilojoules time series
4. **Category_Price_100kJ** — Category-level price per 100kJ averages
5. **Inflation_Volatility** — Composite inflation and volatility measures
6. **Expenditure_Deciles** — Household expenditure by income decile (IES 2022/23)
7. **Rural_CPI** — Rural food CPI series
8. **FAO_Indices** — FAO Food Price Index monthly data
9. **EViews_Master** — Combined panel for EViews econometric modelling
10. **4Way_Classification** — Nutrient-dense/energy-dense × zero-rated/standard-rated
11. **ZR_Premium** — Zero-rated price premium analysis
12. **Methodology** — Data cleaning methodology documentation
13. **Data_Dictionary** — Variable definitions and descriptions
14. **Audit_Log** — Data quality audit trail

## Requirements

```
Python 3.8+
pandas
openpyxl
numpy
```

## Contact

Rochelle van der Walt — rochesvdw@gmail.com
