#!/usr/bin/env python3
"""
Build comprehensive EViews-ready master data file for VAT zero-rating research.
Combines: CPI indices (2008-2025), actual Rand prices (2017-2025), kJ values,
FAO food price indices, and PMBEJD affordability data.
"""

import pandas as pd
import numpy as np
import json
import re
import warnings
warnings.filterwarnings('ignore')

###############################################################################
# 1. LOAD ALL DATA SOURCES
###############################################################################

print("=" * 60)
print("LOADING DATA SOURCES")
print("=" * 60)

# 1a. COICOP 8-digit CPI INDICES (2008-2025, base Dec 2024=100)
df_cpi = pd.read_excel('EXCEL-CPI-COICOP-2018-8digit-202512.xlsx', sheet_name=0)
food_cpi = df_cpi[df_cpi['Division'].astype(str).str.strip() == '1'].copy()
month_cols = [c for c in df_cpi.columns if str(c).startswith('M')]
print(f"  CPI Indices: {len(food_cpi)} food products, {len(month_cols)} months")

# 1b. CPI Average Prices (actual Rand, 2017-2025)
df_prices = pd.read_excel('CPI_Average-Prices_All-urban-202512.xlsx', sheet_name=0)
price_month_cols = [c for c in df_prices.columns if str(c).startswith('M')]
print(f"  Average Prices: {len(df_prices)} rows, {len(price_month_cols)} months")

# 1c. kJ database
with open('kj_per_100g_database.json') as f:
    kj_db = json.load(f)
print(f"  kJ database: {len(kj_db)} items")

# 1d. FAO food price indices
df_fao = pd.read_excel('food_price_indices_data_feb-2.xlsx', 
                        sheet_name='Indices_Monthly_Nominal', header=None, skiprows=4)
df_fao = df_fao.iloc[:, :7]
df_fao.columns = ['Date', 'Food_Price_Index', 'Meat', 'Dairy', 'Cereals', 'Oils', 'Sugar']
df_fao = df_fao.dropna(subset=['Date'])
df_fao['Date'] = pd.to_datetime(df_fao['Date'], errors='coerce')
df_fao = df_fao.dropna(subset=['Date'])
print(f"  FAO Indices: {len(df_fao)} months ({df_fao['Date'].min().date()} to {df_fao['Date'].max().date()})")

# 1e. PMBEJD affordability
df_pmb1 = pd.read_excel('pmbejd_affordability_2025-2.xlsx', sheet_name=0)
df_pmb2 = pd.read_excel('pmbejd_affordability_extended.xlsx', sheet_name=0)
df_pmb = pd.concat([df_pmb1, df_pmb2]).drop_duplicates(subset=['date']).sort_values('date').reset_index(drop=True)
print(f"  PMBEJD: {len(df_pmb)} months ({df_pmb['date'].min().date()} to {df_pmb['date'].max().date()})")

###############################################################################
# 2. UNIT PARSING & PRICE-PER-100g COMPUTATION
###############################################################################

print("\n" + "=" * 60)
print("COMPUTING PRICES PER 100g FROM AVERAGE PRICES")
print("=" * 60)

def parse_unit_grams(unit_str):
    """Parse unit string to grams."""
    if pd.isna(unit_str):
        return None
    s = str(unit_str).strip().lower().replace(',', '.')
    m = re.match(r'([\d.]+)\s*(kilogram|kg|gram|g|litre|liter|l|ml|millilitre)', s)
    if m:
        val = float(m.group(1))
        unit = m.group(2)
        if unit in ['kilogram', 'kg']:
            return val * 1000
        elif unit in ['gram', 'g']:
            return val
        elif unit in ['litre', 'liter', 'l']:
            return val * 1000
        elif unit in ['ml', 'millilitre']:
            return val
    return None

# For each product code, select the SMALLEST standard unit to get the most granular price
# Then compute price per 100g
df_prices['grams'] = df_prices['H08'].apply(parse_unit_grams)
food_prices = df_prices[df_prices['H03'].notna() & df_prices['grams'].notna()].copy()

# For each code, pick the row with the smallest unit
# Group by code and pick smallest grams
price_per_100g = {}
for code, grp in food_prices.groupby('H03'):
    code = int(code)
    # Pick smallest unit row
    smallest = grp.loc[grp['grams'].idxmin()]
    grams = smallest['grams']
    name = smallest['H04']
    
    monthly_prices = {}
    for mc in price_month_cols:
        val = smallest[mc]
        if pd.notna(val) and isinstance(val, (int, float)):
            price_100g = (val / grams) * 100
            monthly_prices[mc] = price_100g
    
    if monthly_prices:
        price_per_100g[code] = {
            'name': name,
            'grams': grams,
            'unit': smallest['H08'],
            'prices': monthly_prices
        }

print(f"  Products with price per 100g: {len(price_per_100g)}")

###############################################################################
# 3. BACK-CALCULATE RAND PRICES FOR 2008-2016 USING CPI INDICES
###############################################################################

print("\n" + "=" * 60)
print("BACK-CALCULATING RAND PRICES (2008-2016)")
print("=" * 60)

# Strategy: CPI index tells us relative price movement.
# If Index_t / Index_ref = Price_t / Price_ref, then:
# Price_t = Price_ref * (Index_t / Index_ref)
# We use Jan 2017 (M201701) as the reference point where we have actual Rand prices.

full_prices_per_100g = {}  # code -> {month_key: price_per_100g}
backfill_count = 0
no_match_count = 0

for _, row in food_cpi.iterrows():
    code = int(row['Eight digit code']) if pd.notna(row['Eight digit code']) else None
    if code is None:
        continue
    name = str(row['Product name']).strip()
    
    # Get CPI index values for all months
    index_values = {}
    for mc in month_cols:
        val = row[mc]
        if pd.notna(val) and isinstance(val, (int, float)) and val > 0:
            index_values[mc] = val
    
    if not index_values:
        continue
    
    # Check if we have actual Rand prices for this product
    if code in price_per_100g:
        actual_prices = price_per_100g[code]['prices']
        
        # Find the best reference month (earliest with both index and actual price)
        ref_month = None
        for mc in sorted(price_month_cols):
            if mc in actual_prices and mc in index_values:
                ref_month = mc
                break
        
        if ref_month:
            ref_price = actual_prices[ref_month]
            ref_index = index_values[ref_month]
            
            # Compute prices for ALL months using index
            computed_prices = {}
            for mc in month_cols:
                if mc in actual_prices:
                    # Use actual price where available
                    computed_prices[mc] = actual_prices[mc]
                elif mc in index_values:
                    # Back-calculate using index ratio
                    computed_prices[mc] = ref_price * (index_values[mc] / ref_index)
                    backfill_count += 1
            
            full_prices_per_100g[code] = {
                'name': name,
                'prices': computed_prices,
                'has_actual': True
            }
        else:
            no_match_count += 1
    else:
        no_match_count += 1

print(f"  Products with full 2008-2025 prices: {len(full_prices_per_100g)}")
print(f"  Back-filled months: {backfill_count}")
print(f"  Products without actual price match: {no_match_count}")

###############################################################################
# 4. CLASSIFY PRODUCTS: CATEGORY + VAT ZERO-RATED STATUS
###############################################################################

print("\n" + "=" * 60)
print("CLASSIFYING PRODUCTS")
print("=" * 60)

# SA VAT zero-rated basic foodstuffs (19 items per Schedule 2 Part B of the VAT Act):
# Brown bread, brown wheaten meal, maize meal, samp, mealie rice,
# dried maize, dried mealies, dried beans, lentils, pilchards/sardines in tins,
# milk powder, dairy powder blend, rice, vegetables, fruit, vegetable oil,
# eggs, edible legumes/pulses, cake flour/bread flour

def classify_product(name, subclass):
    """Classify by food category based on COICOP subclass."""
    if subclass in [1111, 1112, 1113, 1114, 1115]:
        return 'Starchy foods'
    elif subclass in [1161, 1163, 1165, 1169]:
        return 'Fruit & vegetables'
    elif subclass in [1141, 1143, 1145, 1146, 1147, 1148]:
        return 'Dairy & eggs'
    elif subclass in [1122, 1123, 1124, 1125, 1131, 1132, 1133, 1134]:
        return 'Meat, fish & poultry'
    elif subclass in [1151, 1152, 1153]:
        return 'Fats & oils'
    elif subclass in [1171, 1172, 1174, 1175, 1176, 1178, 1179]:
        return 'Sugar & sweets'
    elif subclass in [1181, 1183, 1184, 1185, 1186, 1189]:
        return 'Condiments & other'
    elif subclass in [1191, 1192, 1193, 1194, 1199]:
        return 'Processed foods'
    elif subclass in [1210, 1220, 1230, 1250, 1260, 1290]:
        return 'Beverages'
    return 'Other'

def is_zero_rated(name, subclass, code):
    """Determine if product is VAT zero-rated in SA."""
    nl = name.lower()
    
    # Grains & bread
    if 'brown bread' in nl:
        return True
    if any(x in nl for x in ['maize meal', 'samp', 'mealie rice', 'sorghum meal']):
        return True
    if any(x in nl for x in ['cake flour', 'bread flour']):
        return True
    if 'rice' in nl and subclass == 1111:  # Rice (not rice cakes etc.)
        return True
    
    # Dried legumes
    if any(x in nl for x in ['dried beans', 'lentils', 'dried peas']):
        return True
    
    # Tinned fish
    if any(x in nl for x in ['pilchard', 'sardine']):
        return True
    
    # Dairy
    if any(x in nl for x in ['fresh milk', 'full cream milk', 'low fat milk', 'fat free milk']):
        return True
    if 'milk powder' in nl or 'dairy powder' in nl:
        return True
    
    # Eggs
    if subclass == 1148 or nl.strip() == 'eggs':
        return True
    
    # Vegetable oil/sunflower oil
    if any(x in nl for x in ['vegetable oil', 'sunflower oil', 'cooking oil']):
        return True
    
    # Fresh fruit & vegetables
    if subclass in [1161, 1163, 1165]:
        return True
    
    # Tinned/frozen veg
    if any(x in nl for x in ['tinned vegetable', 'frozen vegetable', 'canned vegetable']):
        return True
    
    return False

# Build full product table
products = []
for _, row in food_cpi.iterrows():
    code = int(row['Eight digit code']) if pd.notna(row['Eight digit code']) else None
    if code is None:
        continue
    name = str(row['Product name']).strip()
    subclass = int(row['Subclass']) if pd.notna(row['Subclass']) else None
    weight = row['Weight'] if pd.notna(row['Weight']) else None
    
    category = classify_product(name, subclass)
    zero_rated = is_zero_rated(name, subclass, code)
    
    # Match kJ value
    kj_val = None
    nl = name.lower()
    if name in kj_db:
        kj_val = kj_db[name]
    else:
        for db_name, db_kj in kj_db.items():
            if db_name.lower() == nl:
                kj_val = db_kj
                break
        if kj_val is None:
            for db_name, db_kj in kj_db.items():
                if db_name.lower() in nl or nl in db_name.lower():
                    kj_val = db_kj
                    break
    
    has_prices = code in full_prices_per_100g
    
    products.append({
        'code': code,
        'name': name,
        'subclass': subclass,
        'cpi_weight': weight,
        'category': category,
        'zero_rated': zero_rated,
        'kj_per_100g': kj_val,
        'has_full_prices': has_prices
    })

df_products = pd.DataFrame(products)
print(f"Total products: {len(df_products)}")
print(f"With kJ values: {df_products['kj_per_100g'].notna().sum()}")
print(f"With full prices: {df_products['has_full_prices'].sum()}")
print(f"Zero-rated: {df_products['zero_rated'].sum()}")

print(f"\nCategory distribution:")
print(df_products['category'].value_counts().to_string())

print(f"\nZero-rated items:")
zr = df_products[df_products['zero_rated']].sort_values('category')
for _, r in zr.iterrows():
    print(f"  [{r['category']}] {r['name']} (code={r['code']}, kJ={r['kj_per_100g']}, prices={r['has_full_prices']})")

###############################################################################
# 5. COMPUTE PRICE PER 100kJ FOR ALL PRODUCTS
###############################################################################

print("\n" + "=" * 60)
print("COMPUTING PRICE PER 100kJ")
print("=" * 60)

# Filter to products that have both kJ values and full prices
eligible = df_products[(df_products['kj_per_100g'].notna()) & (df_products['has_full_prices'])].copy()
print(f"Products with both kJ and prices: {len(eligible)}")

# Build price_per_100kj matrix: products x months
price_per_100kj_data = {}
for _, prod in eligible.iterrows():
    code = prod['code']
    kj = prod['kj_per_100g']
    prices = full_prices_per_100g[code]['prices']
    
    # Price per 100kJ = (price per 100g) / (kJ per 100g) * 100
    price_100kj = {}
    for mc, p100g in prices.items():
        price_100kj[mc] = (p100g / kj) * 100
    
    price_per_100kj_data[code] = price_100kj

###############################################################################
# 6. COMPUTE CATEGORY AVERAGES (weighted by CPI weight)
###############################################################################

print("\n" + "=" * 60)
print("COMPUTING CATEGORY AVERAGES")
print("=" * 60)

categories = ['Starchy foods', 'Fruit & vegetables', 'Dairy & eggs', 
              'Meat, fish & poultry', 'Fats & oils', 'Sugar & sweets',
              'Processed foods']

# Also compute "Zero-rated" as its own line
category_avg = {cat: {} for cat in categories + ['Zero-rated', 'All food']}

for mc in month_cols:
    for cat in categories:
        cat_products = eligible[eligible['category'] == cat]
        values = []
        weights = []
        for _, prod in cat_products.iterrows():
            code = prod['code']
            if code in price_per_100kj_data and mc in price_per_100kj_data[code]:
                values.append(price_per_100kj_data[code][mc])
                w = prod['cpi_weight'] if pd.notna(prod['cpi_weight']) else 1
                weights.append(w)
        if values:
            # Weighted average
            category_avg[cat][mc] = np.average(values, weights=weights)
    
    # Zero-rated average
    zr_products = eligible[eligible['zero_rated']]
    zr_values = []
    zr_weights = []
    for _, prod in zr_products.iterrows():
        code = prod['code']
        if code in price_per_100kj_data and mc in price_per_100kj_data[code]:
            zr_values.append(price_per_100kj_data[code][mc])
            w = prod['cpi_weight'] if pd.notna(prod['cpi_weight']) else 1
            zr_weights.append(w)
    if zr_values:
        category_avg['Zero-rated'][mc] = np.average(zr_values, weights=zr_weights)
    
    # All food average
    all_values = []
    all_weights = []
    for _, prod in eligible.iterrows():
        code = prod['code']
        if code in price_per_100kj_data and mc in price_per_100kj_data[code]:
            all_values.append(price_per_100kj_data[code][mc])
            w = prod['cpi_weight'] if pd.notna(prod['cpi_weight']) else 1
            all_weights.append(w)
    if all_values:
        category_avg['All food'][mc] = np.average(all_values, weights=all_weights)

# Print summary
for cat in categories + ['Zero-rated', 'All food']:
    vals = category_avg[cat]
    if vals:
        first_mc = min(vals.keys())
        last_mc = max(vals.keys())
        print(f"  {cat}: {len(vals)} months, {first_mc}={vals[first_mc]:.4f}, {last_mc}={vals[last_mc]:.4f}")

###############################################################################
# 7. BUILD EVIEWS-READY DATAFRAMES
###############################################################################

print("\n" + "=" * 60)
print("BUILDING EVIEWS-READY DATA")
print("=" * 60)

# Convert month codes to dates
def month_code_to_date(mc):
    """M200801 -> datetime(2008,1,1)"""
    year = int(mc[1:5])
    month = int(mc[5:7])
    return pd.Timestamp(year=year, month=month, day=1)

# Sheet 1: Monthly time series for categories (price per 100kJ)
rows = []
for mc in month_cols:
    dt = month_code_to_date(mc)
    row = {'Date': dt, 'Year': dt.year, 'Month': dt.month}
    for cat in categories + ['Zero-rated', 'All food']:
        col_name = cat.replace(' ', '_').replace(',', '').replace('&', 'and')
        row[col_name] = category_avg[cat].get(mc, np.nan)
    rows.append(row)

df_cat_ts = pd.DataFrame(rows)
print(f"  Category time series: {df_cat_ts.shape}")

# Sheet 2: CPI Price Indices (raw from COICOP 8-digit)
rows2 = []
for mc in month_cols:
    dt = month_code_to_date(mc)
    row = {'Date': dt}
    for _, prod in food_cpi.iterrows():
        code = int(prod['Eight digit code']) if pd.notna(prod['Eight digit code']) else None
        if code:
            val = prod[mc]
            row[f"CPI_{code}"] = val if pd.notna(val) else np.nan
    rows2.append(row)
df_cpi_ts = pd.DataFrame(rows2)
print(f"  CPI indices time series: {df_cpi_ts.shape}")

# Sheet 3: Rand prices per 100g
rows3 = []
for mc in month_cols:
    dt = month_code_to_date(mc)
    row = {'Date': dt}
    for code, data in full_prices_per_100g.items():
        row[f"P100g_{code}"] = data['prices'].get(mc, np.nan)
    rows3.append(row)
df_price_ts = pd.DataFrame(rows3)
print(f"  Price per 100g time series: {df_price_ts.shape}")

# Sheet 4: Price per 100kJ for each product
rows4 = []
for mc in month_cols:
    dt = month_code_to_date(mc)
    row = {'Date': dt}
    for code, prices in price_per_100kj_data.items():
        row[f"PkJ_{code}"] = prices.get(mc, np.nan)
    rows4.append(row)
df_pkj_ts = pd.DataFrame(rows4)
print(f"  Price per 100kJ time series: {df_pkj_ts.shape}")

# Sheet 5: FAO food price indices (2008-2025)
df_fao_08 = df_fao[(df_fao['Date'] >= '2008-01-01') & (df_fao['Date'] <= '2025-12-31')].copy()
df_fao_08 = df_fao_08.reset_index(drop=True)
print(f"  FAO indices (2008-2025): {df_fao_08.shape}")

# Sheet 6: PMBEJD affordability
print(f"  PMBEJD: {df_pmb.shape}")

# Sheet 7: Product lookup table
print(f"  Product lookup: {df_products.shape}")

###############################################################################
# 8. SAVE ALL DATA
###############################################################################

print("\n" + "=" * 60)
print("SAVING INTERMEDIATE DATA")
print("=" * 60)

# Save as pickle for fast reloading
df_cat_ts.to_pickle('_cat_ts.pkl')
df_cpi_ts.to_pickle('_cpi_ts.pkl')
df_price_ts.to_pickle('_price_ts.pkl')
df_pkj_ts.to_pickle('_pkj_ts.pkl')
df_fao_08.to_pickle('_fao_08.pkl')
df_pmb.to_pickle('_pmb.pkl')
df_products.to_pickle('_products.pkl')

# Also save category averages for graphing
with open('_category_averages.json', 'w') as f:
    json.dump({cat: {k: v for k, v in vals.items()} for cat, vals in category_avg.items()}, f)

print("All intermediate data saved.")
print("\nDONE - Data pipeline complete!")
