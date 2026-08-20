"""
generate_ops_data.py
Creates a messy two-week operational export (the kind of raw CSV an ops team
dumps weekly). Intentionally dirty so the pipeline has real cleaning to do.
"""
import numpy as np
import pandas as pd

rng = np.random.default_rng(7)
industries = ["SaaS", "eCommerce", "Healthcare", "Real Estate", "Logistics", "Education", "Finance"]
plans = ["Basic", "Pro", "Enterprise"]
price = {"Basic": 300, "Pro": 900, "Enterprise": 2500}
N = 400  # clients per week

def make_week(week_label, seed_shift):
    r = np.random.default_rng(7 + seed_shift)
    plan = r.choice(plans, N, p=[0.5, 0.35, 0.15])
    mrr = np.array([price[p] for p in plan]) * r.uniform(0.85, 1.3, N)
    df = pd.DataFrame({
        "week": week_label,
        "client_id": [f"CLT-{1000+i}" for i in range(N)],
        "industry": r.choice(industries, N),
        "plan_type": plan,
        "mrr": mrr.round(2),
        "payment_status": r.choice(["Paid", "Late", "Overdue"], N, p=[0.8, 0.13, 0.07]),
        "support_tickets": r.poisson(2, N),
        "active_flag": r.choice([1, 0], N, p=[0.88, 0.12]),
        "new_client": r.choice([1, 0], N, p=[0.06, 0.94]),
        "churned": r.choice([1, 0], N, p=[0.04, 0.96]),
    })
    return df

w1 = make_week("2026-W30", 0)
w2 = make_week("2026-W31", 1)
df = pd.concat([w1, w2], ignore_index=True)

# --- inject realistic mess ---
df["mrr"] = df["mrr"].astype(object)  # allow mixed numeric/string mess
# some mrr values as strings with $ and commas
mask = rng.choice(len(df), 60, replace=False)
df.loc[mask, "mrr"] = df.loc[mask, "mrr"].apply(lambda v: f"${v:,.2f}")
# some blank mrr
df.loc[rng.choice(len(df), 20, replace=False), "mrr"] = ""
# inconsistent casing in plan_type and payment_status
df.loc[rng.choice(len(df), 30, replace=False), "plan_type"] = "enterprise"
df.loc[rng.choice(len(df), 25, replace=False), "payment_status"] = "late"
# blank payment_status
df.loc[rng.choice(len(df), 15, replace=False), "payment_status"] = ""
# a few duplicate rows
dupes = df.sample(10, random_state=1)
df = pd.concat([df, dupes], ignore_index=True)
# whitespace in industry
df.loc[rng.choice(len(df), 20, replace=False), "industry"] = " Healthcare "

df.to_csv("/home/claude/project2/weekly_ops_raw.csv", index=False)
print("Rows (with dupes & mess):", len(df))
print(df.head())
