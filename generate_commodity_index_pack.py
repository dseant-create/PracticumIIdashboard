"""
Generate commodity index tables and cost-vs-commodity analysis from a CPG GFR item table.

Inputs:
  - CPG_GFR_Table_1000.xlsx (generated previously; must include Component 1/2/3 and % columns)

Outputs (in ./cpg_commodity_pack_out):
  - commodity_master.csv
  - commodity_cost_index_monthly.csv
  - item_commodity_bridge.csv
  - item_cost_vs_commodity_trend.csv
  - analysis_scope_column_structure.csv
  - CPG_Commodity_Indices_and_Analysis_Pack.xlsx

Run:
  pip install pandas numpy openpyxl
  python generate_commodity_index_pack.py --items "CPG_GFR_Table_1000.xlsx"
"""

from __future__ import annotations

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd


def commodity_group(name: str) -> str:
    if name in ["Resin", "Packaging", "Fragrance"]:
        return "Petrochemical"
    if name in ["Ingredients", "Raw Materials"]:
        return "Agri & Inputs"
    if name == "Pulp":
        return "Fiber"
    if name == "Active Chemicals":
        return "Industrial Chemicals"
    if name in ["Freight", "Warehousing"]:
        return "Logistics"
    if name in ["Labor", "Co-manufacturing"]:
        return "Conversion"
    if name in ["Utilities", "Overhead"]:
        return "Overhead"
    if name == "Quality Testing":
        return "Quality & Regulatory"
    return "Other"


def base_unit(name: str) -> str:
    if name in ["Resin", "Pulp", "Raw Materials", "Ingredients", "Packaging", "Active Chemicals", "Fragrance"]:
        return "lb"
    if name in ["Freight", "Warehousing"]:
        return "shipment"
    if name in ["Labor", "Co-manufacturing", "Quality Testing"]:
        return "labor hour"
    if name == "Utilities":
        return "kWh"
    if name == "Overhead":
        return "unit"
    return "unit"


def risk_level(name: str) -> str:
    high = {"Resin", "Pulp", "Freight", "Active Chemicals"}
    med = {"Packaging", "Ingredients", "Raw Materials", "Utilities"}
    if name in high:
        return "High"
    if name in med:
        return "Medium"
    return "Low"


def volatility_class(name: str) -> str:
    high = {"Resin", "Freight", "Pulp", "Active Chemicals"}
    mod = {"Packaging", "Ingredients", "Raw Materials", "Utilities"}
    if name in high:
        return "High"
    if name in mod:
        return "Moderate"
    return "Stable"


def drivers(name: str) -> str:
    mapping = {
        "Resin": "Crude oil, natural gas, refinery outages",
        "Packaging": "Resin, paperboard, aluminum, converting capacity",
        "Fragrance": "Petro feedstocks, specialty chemicals, capacity",
        "Ingredients": "Crop yields, dairy/grain prices, seasonality",
        "Raw Materials": "Agricultural inputs, mining, supply constraints",
        "Pulp": "Timber supply, mill capacity, energy costs",
        "Active Chemicals": "Feedstocks, regulatory constraints, plant outages",
        "Freight": "Diesel, capacity, route congestion",
        "Warehousing": "Space rates, labor, utilization",
        "Labor": "Wage inflation, turnover, local market tightness",
        "Co-manufacturing": "Capacity, utilization, contract rates",
        "Utilities": "Electricity/gas prices, regional mix",
        "Overhead": "Fixed cost allocation, depreciation, indirects",
        "Quality Testing": "Lab capacity, reagents, compliance requirements",
    }
    return mapping.get(name, "")


def month_range(start_ym: str, end_ym: str) -> list[pd.Timestamp]:
    start = pd.Period(start_ym, freq="M")
    end = pd.Period(end_ym, freq="M")
    return [p.to_timestamp() for p in pd.period_range(start, end, freq="M")]


def monthly_actual_cost(row: pd.Series, month_starts: list[pd.Timestamp]) -> np.ndarray:
    c23 = float(row["Unit Cost 2023"])
    c24 = float(row["Unit Cost 2024"])
    c25 = float(row["Unit Cost 2025"])
    lcd = pd.to_datetime(row["Last Cost Change Date"])

    changed_24 = round(c24, 4) != round(c23, 4)
    changed_25 = round(c25, 4) != round(c24, 4)

    series = []
    for dt in month_starts:
        y = dt.year
        if y == 2023:
            series.append(c23)
        elif y == 2024:
            if changed_24:
                if lcd.year == 2024 and dt < pd.Timestamp(date(2024, lcd.month, 1)):
                    series.append(c23)
                else:
                    series.append(c24)
            else:
                series.append(c24)
        elif y == 2025:
            if changed_25:
                if lcd.year == 2025 and dt < pd.Timestamp(date(2025, lcd.month, 1)):
                    series.append(c24)
                else:
                    series.append(c25)
            else:
                series.append(c25)
        else:
            series.append(c25)
    return np.array(series, dtype=float)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--items", required=True, help="Path to CPG_GFR_Table_1000.xlsx (or equivalent)")
    parser.add_argument("--out", default="cpg_commodity_pack_out", help="Output directory")
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    random.seed(args.seed)
    np.random.seed(args.seed)

    items_path = Path(args.items)
    if not items_path.exists():
        raise FileNotFoundError(f"Input file not found: {items_path}")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    items = pd.read_excel(items_path, sheet_name=0)
    items["UPC"] = items["UPC"].astype(str)

    comp_cols = ["Component 1", "Component 2", "Component 3"]
    unique_components = sorted(set(pd.concat([items[c].astype(str) for c in comp_cols], ignore_index=True).unique().tolist()))

    commodity_master = pd.DataFrame(
        [
            {
                "Commodity_ID": f"COM-{i+1:03d}",
                "Commodity_Name": name,
                "Component_Type": name,
                "Commodity_Group": commodity_group(name),
                "Base_Unit": base_unit(name),
                "Currency": "USD",
                "Source_Type": "Modeled Index",
                "Key_Drivers": drivers(name),
                "Risk_Level": risk_level(name),
                "Volatility_Class": volatility_class(name),
                "Hedging_Eligible": "Y" if name in {"Resin", "Pulp", "Freight", "Active Chemicals"} else "N",
                "Notes": "",
            }
            for i, name in enumerate(unique_components)
        ]
    )

    # Monthly indices 2023–2025 (36 months)
    months = month_range("2023-01", "2025-12")
    n_months = len(months)

    # Commodity-specific drift/vol assumptions (monthly)
    assumptions = {}
    for _, row in commodity_master.iterrows():
        name = row["Commodity_Name"]
        if name in {"Resin", "Freight", "Pulp", "Active Chemicals"}:
            drift = random.uniform(0.000, 0.004)
            vol = random.uniform(0.015, 0.040)
        elif name in {"Packaging", "Ingredients", "Raw Materials", "Utilities"}:
            drift = random.uniform(0.001, 0.005)
            vol = random.uniform(0.008, 0.020)
        else:
            drift = random.uniform(0.001, 0.004)
            vol = random.uniform(0.004, 0.012)
        assumptions[name] = (drift, vol)

    base_unit_cost = {
        "Resin": 1.55,
        "Pulp": 1.05,
        "Packaging": 0.95,
        "Ingredients": 1.20,
        "Raw Materials": 1.10,
        "Active Chemicals": 1.80,
        "Fragrance": 3.50,
        "Freight": 18.00,
        "Warehousing": 6.50,
        "Labor": 24.00,
        "Co-manufacturing": 20.00,
        "Utilities": 0.14,
        "Overhead": 0.40,
        "Quality Testing": 35.00,
    }

    index_rows = []
    for _, cm in commodity_master.iterrows():
        name = cm["Commodity_Name"]
        drift, vol = assumptions[name]

        idx = [100.0]
        for t in range(1, n_months):
            month_num = months[t].month
            seasonal = 0.0
            if name in {"Freight", "Warehousing"}:
                seasonal = 0.004 if month_num in {10, 11, 12} else (-0.002 if month_num in {1, 2} else 0.0)
            if name in {"Ingredients", "Raw Materials"}:
                seasonal = 0.003 if month_num in {8, 9} else 0.0
            shock = np.random.normal(drift + seasonal, vol)
            idx.append(max(60.0, idx[-1] * (1 + shock)))

        idx = np.array(idx)
        base = base_unit_cost.get(name, 1.0)
        unit_cost = base * (idx / 100.0)

        mom = np.concatenate([[np.nan], (idx[1:] / idx[:-1] - 1)])
        yoy = np.array([np.nan] * n_months)
        for t in range(12, n_months):
            yoy[t] = idx[t] / idx[t - 12] - 1
        vol_score = pd.Series(mom).rolling(6).std().values

        for t, dt in enumerate(months):
            index_rows.append(
                {
                    "Commodity_ID": cm["Commodity_ID"],
                    "Commodity_Name": name,
                    "Commodity_Group": cm["Commodity_Group"],
                    "Component_Type": cm["Component_Type"],
                    "Base_Unit": cm["Base_Unit"],
                    "Currency": cm["Currency"],
                    "Source_Type": cm["Source_Type"],
                    "Year": dt.year,
                    "Month": dt.month,
                    "Month_Start": dt.date().isoformat(),
                    "Cost_Index": round(float(idx[t]), 2),
                    "Unit_Cost": round(float(unit_cost[t]), 4),
                    "MoM_Change_%": None if np.isnan(mom[t]) else round(float(mom[t]), 5),
                    "YoY_Change_%": None if np.isnan(yoy[t]) else round(float(yoy[t]), 5),
                    "Volatility_Score_6m": None if np.isnan(vol_score[t]) else round(float(vol_score[t]), 5),
                    "Notes": "",
                }
            )

    commodity_index_monthly = pd.DataFrame(index_rows)
    commodity_index_monthly["Month_Start"] = pd.to_datetime(commodity_index_monthly["Month_Start"])
    month_starts = sorted(commodity_index_monthly["Month_Start"].unique())

    # Baseline (Jan 2023) index per commodity
    jan2023 = pd.Timestamp("2023-01-01")
    baseline_idx = (
        commodity_index_monthly[commodity_index_monthly["Month_Start"] == jan2023]
        .set_index("Commodity_ID")["Cost_Index"]
        .to_dict()
    )

    idx_by_month = {}
    for dt, sub in commodity_index_monthly.groupby("Month_Start"):
        idx_by_month[dt] = sub.set_index("Commodity_ID")["Cost_Index"].to_dict()

    comp_to_id = dict(zip(commodity_master["Component_Type"], commodity_master["Commodity_ID"]))

    # Bridge
    bridge_rows = []
    for _, r in items.iterrows():
        upc = r["UPC"]
        supplier_id = r.get("Supplier ID", None)
        region = r.get("Region #", None)
        for j in [1, 2, 3]:
            comp = r[f"Component {j}"]
            pct = float(r[f"Component {j} %"])
            bridge_rows.append(
                {
                    "UPC": upc,
                    "Commodity_ID": comp_to_id.get(comp),
                    "Component_Type": comp,
                    "Component_%": pct,
                    "Primary_Flag": "Y" if j == 1 else "N",
                    "Supplier_ID": supplier_id,
                    "Region_#": region,
                }
            )

    item_commodity_bridge = pd.DataFrame(bridge_rows)

    comp_weights = item_commodity_bridge.groupby("UPC").apply(
        lambda g: list(zip(g["Commodity_ID"].tolist(), g["Component_%"].astype(float).tolist()))
    )

    # Trend table
    trend_rows = []
    for _, r in items.iterrows():
        upc = r["UPC"]
        weights = comp_weights.loc[upc]
        c23_base = float(r["Unit Cost 2023"])
        actual_cost_series = monthly_actual_cost(r, month_starts)

        for idx_m, dt in enumerate(month_starts):
            year = dt.year
            weighted_index = 0.0
            expected_multiplier = 0.0
            for com_id, w in weights:
                idx_level = idx_by_month[dt].get(com_id, baseline_idx.get(com_id, 100.0))
                weighted_index += w * idx_level
                expected_multiplier += w * (idx_level / baseline_idx.get(com_id, 100.0))

            expected_cost = c23_base * expected_multiplier
            actual_cost = float(actual_cost_series[idx_m])

            price = float(r[f"Sale Price {year}"])
            baseline_margin = float(r[f"Unit Margin $ {year}"])
            actual_margin = price - actual_cost
            margin_impact = actual_margin - baseline_margin

            cost_var = actual_cost - expected_cost
            cost_var_pct = cost_var / expected_cost if expected_cost != 0 else np.nan
            supplier_effect = "Supplier" if abs(cost_var_pct) > 0.02 else "Market"

            trend_rows.append(
                {
                    "UPC": upc,
                    "Month_Start": dt.date().isoformat(),
                    "Year": year,
                    "Month": dt.month,
                    "Weighted_Commodity_Index": round(weighted_index, 2),
                    "Expected_Cost": round(expected_cost, 4),
                    "Actual_Item_Cost": round(actual_cost, 4),
                    "Cost_Variance_$": round(cost_var, 4),
                    "Cost_Variance_%": None if np.isnan(cost_var_pct) else round(float(cost_var_pct), 5),
                    "Sale_Price_Proxy": round(price, 2),
                    "Baseline_Unit_Margin": round(baseline_margin, 4),
                    "Actual_Unit_Margin": round(actual_margin, 4),
                    "Margin_Impact_$": round(margin_impact, 4),
                    "Supplier_Effect_Flag": supplier_effect,
                }
            )

    trend = pd.DataFrame(trend_rows)
    trend["Month_Start"] = pd.to_datetime(trend["Month_Start"])
    trend = trend.sort_values(["UPC", "Month_Start"])
    trend["Expected_Cost_YoY_%"] = trend.groupby("UPC")["Expected_Cost"].pct_change(12)
    trend["Actual_Cost_YoY_%"] = trend.groupby("UPC")["Actual_Item_Cost"].pct_change(12)

    is_supplier = (trend["Supplier_Effect_Flag"] == "Supplier").astype(int)
    trend["Supplier_Streak"] = (
        is_supplier.groupby(trend["UPC"])
        .apply(lambda s: s.groupby((s != s.shift()).cumsum()).cumsum())
        .reset_index(level=0, drop=True)
    )
    trend["Negotiation_Flag"] = np.where(
        ((trend["Expected_Cost_YoY_%"] < -0.01) & (trend["Actual_Cost_YoY_%"] > 0.01)) | (trend["Supplier_Streak"] >= 3),
        "Y",
        "N",
    )

    # Price lag months heuristic (annual)
    def sign(x: float, eps: float = 1e-9) -> int:
        if x > eps:
            return 1
        if x < -eps:
            return -1
        return 0

    lag_rows = []
    for _, r in items.iterrows():
        upc = r["UPC"]
        p23, p24, p25 = float(r["Sale Price 2023"]), float(r["Sale Price 2024"]), float(r["Sale Price 2025"])
        c23, c24, c25 = float(r["Unit Cost 2023"]), float(r["Unit Cost 2024"]), float(r["Unit Cost 2025"])
        dp24, dc24 = sign(p24 - p23), sign(c24 - c23)
        dp25, dc25 = sign(p25 - p24), sign(c25 - c24)
        lag_24 = 0 if dp24 == dc24 else (6 if dp25 == dc24 else 3)
        lag_25 = 0 if dp25 == dc25 else 3
        lag_rows.append({"UPC": upc, "Price_Lag_Months_2024": lag_24, "Price_Lag_Months_2025": lag_25})
    price_lag = pd.DataFrame(lag_rows)

    trend = trend.merge(price_lag, on="UPC", how="left")

    analysis_scope = pd.DataFrame(
        [
            {
                "Layer": "commodity_master",
                "Scope": "Commodity dimension: definitions, units, risk/volatility tags, drivers, hedging.",
                "Grain": "One row per commodity component type.",
                "Key Headings": "Commodity_ID, Commodity_Name, Component_Type, Commodity_Group, Base_Unit, Currency, Source_Type, Key_Drivers, Risk_Level, Volatility_Class, Hedging_Eligible",
            },
            {
                "Layer": "commodity_cost_index_monthly",
                "Scope": "Monthly commodity cost index and unit cost trend, normalized to Jan-2023=100.",
                "Grain": "Commodity_ID x Month (36 months for 2023–2025).",
                "Key Headings": "Commodity_ID, Month_Start, Year, Month, Cost_Index, Unit_Cost, MoM_Change_%, YoY_Change_%, Volatility_Score_6m, Notes",
            },
            {
                "Layer": "item_commodity_bridge",
                "Scope": "Mapping of items to their 3 cost components with weights.",
                "Grain": "UPC x Component (3 rows per UPC).",
                "Key Headings": "UPC, Commodity_ID, Component_Type, Component_%, Primary_Flag, Supplier_ID, Region_#",
            },
            {
                "Layer": "item_cost_vs_commodity_trend",
                "Scope": "Monthly expected cost modeled from commodity indices vs actual item cost, with variance and margin impact.",
                "Grain": "UPC x Month (36 rows per UPC).",
                "Key Headings": "UPC, Month_Start, Weighted_Commodity_Index, Expected_Cost, Actual_Item_Cost, Cost_Variance_$, Cost_Variance_%, Sale_Price_Proxy, Baseline_Unit_Margin, Actual_Unit_Margin, Margin_Impact_$, Supplier_Effect_Flag, Expected_Cost_YoY_%, Actual_Cost_YoY_%, Negotiation_Flag, Price_Lag_Months_2024/2025",
            },
        ]
    )

    # Write outputs
    commodity_master.to_csv(out_dir / "commodity_master.csv", index=False)
    commodity_index_monthly.to_csv(out_dir / "commodity_cost_index_monthly.csv", index=False)
    item_commodity_bridge.to_csv(out_dir / "item_commodity_bridge.csv", index=False)
    trend.to_csv(out_dir / "item_cost_vs_commodity_trend.csv", index=False)
    analysis_scope.to_csv(out_dir / "analysis_scope_column_structure.csv", index=False)

    xlsx_path = out_dir / "CPG_Commodity_Indices_and_Analysis_Pack.xlsx"
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        commodity_master.to_excel(writer, sheet_name="commodity_master", index=False)
        commodity_index_monthly.to_excel(writer, sheet_name="commodity_index_monthly", index=False)
        item_commodity_bridge.to_excel(writer, sheet_name="item_commodity_bridge", index=False)
        trend.to_excel(writer, sheet_name="item_cost_vs_commodity_trend", index=False)
        analysis_scope.to_excel(writer, sheet_name="analysis_scope", index=False)

    print(f"Done. Outputs in: {out_dir.resolve()}")


if __name__ == "__main__":
    main()
