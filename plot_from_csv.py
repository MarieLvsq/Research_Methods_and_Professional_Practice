import pandas as pd
import matplotlib.pyplot as plt

# ===============================
# Load data
# ===============================
superplus  = pd.read_excel("Superplus.xlsx")
heather    = pd.read_excel("Heather.xlsx")
diets      = pd.read_excel("Diets.xlsx")
designs    = pd.read_excel("Designs.xlsx")
brandprefs = pd.read_excel("Brandprefs.xlsx")

# ===============================
# SUPERPLUS (Data Set C): Income by Sex
# ===============================
# Summary -> CSV
superplus_summary = superplus.groupby("Sex")["Income"].agg(
    Count="count", Mean="mean", Median="median", SD="std", Min="min", Max="max"
).round(3)
superplus_summary.to_csv("Superplus_Summary.csv")

# Histograms by sex (two separate charts)
for sex in ["M", "F"]:
    data = superplus.loc[superplus["Sex"] == sex, "Income"]
    plt.figure(figsize=(7,5))
    plt.hist(data, bins=12)  # default style/colors
    plt.title(f"Superplus Income Histogram – {sex}")
    plt.xlabel("Income (£'000s)")
    plt.ylabel("Frequency")
    plt.tight_layout()
    plt.savefig(f"Superplus_Hist_{sex}.png", dpi=120)
    plt.close()

# Optional grouped bar of binned percentages by sex
bin_edges = list(range(int(superplus["Income"].min()//10*10), int(superplus["Income"].max()//10*10 + 20), 10))
labels = [f"{bin_edges[i]}–{bin_edges[i+1]}" for i in range(len(bin_edges)-1)]

def pct_by_bin(df, value_col, group_col, group_value, edges):
    sub = df[df[group_col] == group_value][value_col]
    counts, _ = pd.cut(sub, bins=edges, right=True, include_lowest=True, retbins=True)
    counts = counts.value_counts(sort=False)
    pct = (counts / counts.sum()).reindex(counts.index, fill_value=0).values
    return pct

m_pct = pct_by_bin(superplus, "Income", "Sex", "M", bin_edges)
f_pct = pct_by_bin(superplus, "Income", "Sex", "F", bin_edges)

x = range(len(labels))
width = 0.45
plt.figure(figsize=(9,5))
plt.bar([i - width/2 for i in x], m_pct, width, label="Male")
plt.bar([i + width/2 for i in x], f_pct, width, label="Female")
plt.xticks(x, labels, rotation=45)
plt.ylabel("Percentage")
plt.title("Superplus Income Distribution by Sex (Binned %)")
plt.legend()
plt.tight_layout()
plt.savefig("Superplus_GroupedBar.png", dpi=120)
plt.close()

# ===============================
# HEATHER (Data Set E): Prevalence by Location
# ===============================
heather_clean = heather.dropna().rename(
    columns={"Frequencies":"Prevalence", "Unnamed: 1":"Location A", "Unnamed: 2":"Location B"}
)
heather_clean = heather_clean[heather_clean["Prevalence"].isin(["Absent","Sparse","Abundant"])]
heather_clean[["Location A","Location B"]] = heather_clean[["Location A","Location B"]].astype(int)
tot_A = heather_clean["Location A"].sum()
tot_B = heather_clean["Location B"].sum()
heather_pct = heather_clean.copy()
heather_pct["Pct_A"] = (heather_pct["Location A"] / tot_A * 100).round(1)
heather_pct["Pct_B"] = (heather_pct["Location B"] / tot_B * 100).round(1)
heather_pct.to_csv("Heather_Percentages.csv", index=False)

# Clustered column (A vs B by Prevalence)
cats = heather_pct["Prevalence"].tolist()
A_vals = heather_pct["Pct_A"].tolist()
B_vals = heather_pct["Pct_B"].tolist()
x = range(len(cats))
width = 0.4
plt.figure(figsize=(7,5))
plt.bar([i - width/2 for i in x], A_vals, width, label="Location A")
plt.bar([i + width/2 for i in x], B_vals, width, label="Location B")
plt.xticks(x, cats)
plt.ylabel("Percentage")
plt.title("Heather Prevalence by Location")
plt.legend()
plt.tight_layout()
plt.savefig("Heather_Clustered.png", dpi=120)
plt.close()

# ===============================
# DIETS (Data Set B): Weight Loss
# ===============================
diet_summary = diets.groupby("Diet")["Wtloss"].agg(
    Count="count", Mean="mean", Median="median", SD="std", Min="min", Max="max"
).round(3)
diet_summary.to_csv("Diet_Summaries.csv")

# Relative frequency histograms with bin width = 2
bins = list(range(-6, 14, 2))  # (-6,-4],..., (12,14]
diets["Class"] = pd.cut(diets["Wtloss"], bins=bins, right=True, include_lowest=False)

diet_hist = diets.groupby(["Diet", "Class"]).size().reset_index(name="Frequency")
diet_hist["Relative_Freq"] = diet_hist.groupby("Diet")["Frequency"].transform(lambda x: x / x.sum())
diet_hist["Class"] = diet_hist["Class"].astype(str)
diet_hist = diet_hist.sort_values(["Diet","Class"]).reset_index(drop=True)
diet_hist["Relative_Freq"] = diet_hist["Relative_Freq"].round(4)
diet_hist.to_csv("Diet_Histogram.csv", index=False)

# Plot Diet A histogram (relative frequency)
dietA = diets[diets["Diet"]=="A"]["Wtloss"]
plt.figure(figsize=(7,5))
plt.hist(dietA, bins=bins)
plt.title("Diet A – Relative Frequency Histogram")
plt.xlabel("Weight Loss (kg)")
plt.ylabel("Frequency")
plt.tight_layout()
plt.savefig("DietA_Hist.png", dpi=120)
plt.close()

# Plot Diet B histogram (relative frequency)
dietB = diets[diets["Diet"]=="B"]["Wtloss"]
plt.figure(figsize=(7,5))
plt.hist(dietB, bins=bins)
plt.title("Diet B – Relative Frequency Histogram")
plt.xlabel("Weight Loss (kg)")
plt.ylabel("Frequency")
plt.tight_layout()
plt.savefig("DietB_Hist.png", dpi=120)
plt.close()

# ===============================
# DESIGNS (Data Set F): Con1 vs Con2 by Store
# ===============================
designs_summary = designs[["Con1","Con2"]].agg(["mean","median","std","min","max"]).round(3)
designs_summary.to_csv("Designs_Summary.csv")

designs_diff = designs.assign(Diff=designs["Con1"] - designs["Con2"])
designs_diff.to_csv("Designs_Differences.csv", index=False)

# Clustered bar per store
x = range(len(designs))
width = 0.4
plt.figure(figsize=(9,5))
plt.bar([i - width/2 for i in x], designs["Con1"], width, label="Con1")
plt.bar([i + width/2 for i in x], designs["Con2"], width, label="Con2")
plt.xticks(list(x), designs["Store"])
plt.xlabel("Store")
plt.ylabel("Units Sold")
plt.title("Designs – Sales by Store (Con1 vs Con2)")
plt.legend()
plt.tight_layout()
plt.savefig("Designs_ByStore.png", dpi=120)
plt.close()

# Mean comparison bar
means = designs[["Con1","Con2"]].mean()
plt.figure(figsize=(6,5))
plt.bar(["Con1","Con2"], means.values)
plt.ylabel("Mean Units Sold")
plt.title("Designs – Mean Sales Comparison")
plt.tight_layout()
plt.savefig("Designs_Mean.png", dpi=120)
plt.close()

# ===============================
# BRANDPREFS (Data Set D): Brand by Area
# ===============================
brand_counts = brandprefs.groupby(["Area","Brand"]).size().reset_index(name="Count")
area_totals = brand_counts.groupby("Area")["Count"].sum().reset_index(name="Total")
brand_pct = brand_counts.merge(area_totals, on="Area")
brand_pct["Pct"] = (brand_pct["Count"] / brand_pct["Total"] * 100).round(1)
brand_pct.to_csv("BrandPrefs_Percentages.csv", index=False)

# Separate bar charts for Area 1 and Area 2
for area in sorted(brand_pct["Area"].unique()):
    df = brand_pct[brand_pct["Area"] == area].sort_values("Brand")
    plt.figure(figsize=(6,5))
    plt.bar(df["Brand"], df["Pct"])
    plt.ylim(0, 100)
    plt.ylabel("Percentage")
    plt.title(f"Brand Preferences – Area {area}")
    plt.tight_layout()
    plt.savefig(f"Brandprefs_Area{area}.png", dpi=120)
    plt.close()

# Clustered comparison (A,B,Other) across areas
brands = ["A","B","Other"]
area_vals = {area: brand_pct[brand_pct["Area"]==area].set_index("Brand").reindex(brands)["Pct"].values
             for area in sorted(brand_pct["Area"].unique())}
x = range(len(brands))
width = 0.35
plt.figure(figsize=(7,5))
offsets = [-width/2, width/2]
for idx, area in enumerate(sorted(area_vals.keys())):
    plt.bar([i + offsets[idx] for i in x], area_vals[area], width, label=f"Area {area}")
plt.xticks(list(x), brands)
plt.ylabel("Percentage")
plt.title("Brand Preferences – Area 1 vs Area 2")
plt.legend()
plt.tight_layout()
plt.savefig("Brandprefs_Clustered.png", dpi=120)
plt.close()

print("Done. CSVs and PNGs saved in the current folder.")
