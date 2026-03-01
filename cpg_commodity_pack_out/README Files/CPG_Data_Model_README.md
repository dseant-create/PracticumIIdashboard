# CPG Commodity Cost Modeling & Analytics Pack

## Overview

This project simulates a Consumer Packaged Goods (CPG) sourcing and cost
analytics environment.\
It includes:

-   A 1,000-item synthetic GFR (Goods for Resale) dataset
-   Commodity master data
-   Monthly commodity cost indices (2023--2025)
-   Item-to-commodity mapping bridge
-   Expected vs actual cost modeling
-   Margin and negotiation analytics framework
-   Reproducible Python data generators

This repository is designed for: - CPG sourcing analytics practice -
Cost modeling simulation - Margin exposure analysis - Supplier
negotiation scenario modeling - BI / dashboard development (Power BI,
Tableau, etc.)

------------------------------------------------------------------------

# Project Structure

## 1. Item Dataset

**File:** `CPG_GFR_Table_1000.xlsx`

Simulated 1,000 SKU dataset including: - Item Category: GFR - Department
(1--20) - UPC - Supplier details - Store and region data - 2023--2025
sale price, units, and total sales - Margin % and unit margin - Unit
cost by year - Last cost change date - 3-component cost breakdown per
item

Cost components sum to 100% and roll into total unit cost.

------------------------------------------------------------------------

## 2. Commodity Master Table

**File:** `commodity_master.csv`

One row per component commodity including:

-   Commodity_ID
-   Commodity_Name
-   Component_Type
-   Commodity_Group
-   Base_Unit
-   Risk_Level
-   Volatility_Class
-   Hedging_Eligible
-   Key_Drivers

Used as the dimensional backbone of the cost model.

------------------------------------------------------------------------

## 3. Commodity Cost Index (Monthly)

**File:** `commodity_cost_index_monthly.csv`

Monthly index data for 2023--2025 (36 months per commodity):

-   Cost_Index (Jan-2023 = 100 baseline)
-   Unit_Cost (scaled)
-   MoM_Change\_%
-   YoY_Change\_%
-   Volatility_Score_6m

Modeled using stochastic drift + volatility assumptions by commodity
group.

------------------------------------------------------------------------

## 4. Item--Commodity Bridge

**File:** `item_commodity_bridge.csv`

Mapping of each SKU to its 3 cost components:

-   UPC
-   Commodity_ID
-   Component\_%
-   Primary_Flag
-   Supplier_ID
-   Region\_#

Enables weighted commodity exposure modeling.

------------------------------------------------------------------------

## 5. Item Cost vs Commodity Trend

**File:** `item_cost_vs_commodity_trend.csv`

Monthly comparison model (36 months × 1,000 SKUs):

-   Weighted_Commodity_Index
-   Expected_Cost (commodity-modeled)
-   Actual_Item_Cost
-   Cost_Variance\_\$
-   Cost_Variance\_%
-   Margin_Impact\_\$
-   Supplier_Effect_Flag
-   Negotiation_Flag
-   Price_Lag_Months

Supports analysis of: - Commodity pass-through - Supplier-driven cost
variance - Margin exposure - Pricing lag behavior - Negotiation
candidates

------------------------------------------------------------------------

# Analytical Use Cases

## Commodity Pass-Through Analysis

Compare: - Expected cost (commodity-driven) - Actual cost
(supplier-driven)

Identify supplier inefficiencies or lagging cost reductions.

## Margin Risk Monitoring

Evaluate: - Margin erosion due to commodity spikes - Retail pricing lag
impact - Category-level exposure

## Negotiation Targeting

Flag SKUs where: - Commodity index down YoY - Actual cost up YoY -
Supplier variance streak \> threshold

------------------------------------------------------------------------

# Dashboarding Recommendations

### Executive View

-   Avg Cost Variance %
-   \% SKUs flagged for negotiation
-   Total Margin Impact
-   Top volatile commodities

### Commodity Deep Dive

-   Monthly index trends
-   YoY heatmap
-   Volatility vs risk classification

### Supplier Performance

-   Avg cost variance by supplier
-   Negotiation flags by supplier
-   Margin exposure by region

------------------------------------------------------------------------

# Technical Requirements

Python 3.10+

Install dependencies: pip install pandas numpy openpyxl pypandoc

To regenerate the commodity pack: python
generate_commodity_index_pack.py --items CPG_GFR_Table_1000.xlsx

------------------------------------------------------------------------

# Data Model Summary

Dimension Tables: - commodity_master - (optional) calendar dimension -
supplier/store dimension (extendable)

Fact Tables: - commodity_cost_index_monthly -
item_cost_vs_commodity_trend - item_commodity_bridge

------------------------------------------------------------------------

# Notes

All data is synthetic and modeled for analytical practice only.\
Indices are simulated using stochastic processes and do not represent
real commodity markets.

------------------------------------------------------------------------

# End of README
