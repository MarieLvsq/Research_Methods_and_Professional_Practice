import pandas as pd

# ---------- SUPERPLUS ----------
superplus = pd.read_excel("Superplus.xlsx")
superplus_summary = superplus.groupby("Sex")["Income"].agg(
    Count="count", Mean="mean", Median="median", SD="std", Min="min", Max="max"
).round(3)
superplus_summary.to_csv("Superplus_Summary.csv")

# ---------- HEATHER ----------
heather = pd.read_excel("Heather.xlsx")
heather_clean = heather.dropna().rename(
    columns={"Frequencies": "Prevalence", "Unnamed: 1": "Location A", "Unnamed: 2": "Location B"}
)
heather_clean = heather_clean[heather_clean["Prevalence"].isin(["Absent", "Sparse", "Abundant"])]
heather_clean[["Location A", "Location B"]] = heather_clean[["Location A", "Location B"]].astype(int)

tot_A = heather_clean["Location A"].sum()
tot_B = heather_clean["Location B"].sum()
heather_pct = heather_clean.copy()
heather_pct["Pct_A"] = (heather_clean["Location A"] / tot_A * 100).round(1)
heather_pct["Pct_B"] = (heather_clean["Location B"] / tot_B * 100).round(1)

heather_pct.to_csv("Heather_Percentages.csv", index=False)

# ---------- DIETS ----------
diets = pd.read_excel("Diets.xlsx")
diet_summary = diets.groupby("Diet")["Wtloss"].agg(
    Count="count", Mean="mean", Median="median", SD="std", Min="min", Max="max"
).round(3)
diet_summary.to_csv("Diet_Summaries.csv")

# Histogram bins
bins = list(range(-6, 14, 2))
diets["Class"] = pd.cut(diets["Wtloss"], bins=bins, right=True, include_lowest=False)
diet_hist = diets.groupby(["Diet", "Class"]).size().reset_index(name="Frequency")
diet_hist["Relative_Freq"] = diet_hist.groupby("Diet")["Frequency"].transform(lambda x: x / x.sum())
diet_hist["Class"] = diet_hist["Class"].astype(str)
diet_hist = diet_hist.sort_values(["Diet", "Class"]).reset_index(drop=True)
diet_hist["Relative_Freq"] = diet_hist["Relative_Freq"].round(4)
diet_hist.to_csv("Diet_Histogram.csv", index=False)

# ---------- DESIGNS ----------
designs = pd.read_excel("Designs.xlsx")
designs_summary = designs[["Con1", "Con2"]].agg(["mean", "median", "std", "min", "max"]).round(3)
designs_summary.to_csv("Designs_Summary.csv")

designs_diff = designs.assign(Diff=designs["Con1"] - designs["Con2"])
designs_diff.to_csv("Designs_Differences.csv", index=False)

# ---------- BRANDPREFS ----------
brandprefs = pd.read_excel("Brandprefs.xlsx")
brand_counts = brandprefs.groupby(["Area", "Brand"]).size().reset_index(name="Count")
area_totals = brand_counts.groupby("Area")["Count"].sum().reset_index(name="Total")
brand_pct = brand_counts.merge(area_totals, on="Area")
brand_pct["Pct"] = (brand_pct["Count"] / brand_pct["Total"] * 100).round(1)

brand_pct.to_csv("BrandPrefs_Percentages.csv", index=False)
