import pandas as pd
import os

# ---------- Paths ----------
BASE = "/mnt/data"

# ---------- Load datasets ----------
superplus  = pd.read_excel(os.path.join(BASE, "Superplus.xlsx"))
heather    = pd.read_excel(os.path.join(BASE, "Heather.xlsx"))
diets      = pd.read_excel(os.path.join(BASE, "Diets.xlsx"))
designs    = pd.read_excel(os.path.join(BASE, "Designs.xlsx"))
brandprefs = pd.read_excel(os.path.join(BASE, "Brandprefs.xlsx"))

# ---------- SUPERPLUS: summary by sex ----------
superplus_summary = superplus.groupby("Sex")["Income"].agg(
    Count="count", Mean="mean", Median="median", SD="std", Min="min", Max="max"
).round(3)
superplus_summary.to_csv(os.path.join(BASE, "Superplus_Summary.csv"))

# ---------- HEATHER: prevalence percentages ----------
# Clean sheet headers into tidy table
heather_clean = heather.dropna().rename(
    columns={
        "Frequencies": "Prevalence",
        "Unnamed: 1": "Location A",
        "Unnamed: 2": "Location B",
    }
)
heather_clean = heather_clean[heather_clean["Prevalence"].isin(["Absent", "Sparse", "Abundant"])]
heather_clean[["Location A", "Location B"]] = heather_clean[["Location A", "Location B"]].astype(int)

# Totals per location to compute %
tot_A = heather_clean["Location A"].sum()
tot_B = heather_clean["Location B"].sum()

heather_pct = heather_clean.copy()
heather_pct["Pct_A"] = (heather_pct["Location A"] / tot_A * 100).astype(float).round(1)
heather_pct["Pct_B"] = (heather_pct["Location B"] / tot_B * 100).astype(float).round(1)

heather_pct.to_csv(os.path.join(BASE, "Heather_Percentages.csv"), index=False)

# ---------- DIETS: summaries & relative-frequency histogram ----------
diet_summary = diets.groupby("Diet")["Wtloss"].agg(
    Count="count", Mean="mean", Median="median", SD="std", Min="min", Max="max"
).round(3)
diet_summary.to_csv(os.path.join(BASE, "Diet_Summaries.csv"))

# Histogram bins (width=2kg), same scheme used in worksheet
bins = list(range(-6, 14, 2))  # (-6,-4],(-4,-2],..., (12,14] (top bins may be empty)
diets["Class"] = pd.cut(diets["Wtloss"], bins=bins, right=True, include_lowest=False)

diet_hist = diets.groupby(["Diet", "Class"]).size().reset_index(name="Frequency")
diet_hist["Relative_Freq"] = diet_hist.groupby("Diet")["Frequency"].transform(lambda x: x / x.sum())
diet_hist["Class"] = diet_hist["Class"].astype(str)
diet_hist = diet_hist.sort_values(["Diet", "Class"]).reset_index(drop=True)
diet_hist["Relative_Freq"] = diet_hist["Relative_Freq"].round(4)

diet_hist.to_csv(os.path.join(BASE, "Diet_Histogram.csv"), index=False)

# ---------- DESIGNS: summary stats & per-store differences ----------
designs_summary = designs[["Con1", "Con2"]].agg(["mean", "median", "std", "min", "max"]).round(3)
designs_summary.to_csv(os.path.join(BASE, "Designs_Summary.csv"))

designs_diff = designs.assign(Diff=designs["Con1"] - designs["Con2"])
designs_diff.to_csv(os.path.join(BASE, "Designs_Differences.csv"), index=False)

# ---------- BRANDPREFS: percentage by area & brand ----------
brand_counts = brandprefs.groupby(["Area", "Brand"]).size().reset_index(name="Count")
area_totals  = brand_counts.groupby("Area")["Count"].sum().reset_index(name="Total")
brand_pct    = brand_counts.merge(area_totals, on="Area")
brand_pct["Pct"] = (brand_pct["Count"] / brand_pct["Total"] * 100).astype(float).round(1)

brand_pct.to_csv(os.path.join(BASE, "BrandPrefs_Percentages.csv"), index=False)
