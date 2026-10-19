"""
A/B analysis of outreach variants from a CSV log (same columns as the Outreach_Log sheet).
Run: python ab_test.py [path/to/log.csv]
Default file data/outreach_log_SAMPLE.csv is SIMULATED - replace with your real log before quoting any numbers.
"""
import sys
import numpy as np
import pandas as pd
from scipy.stats import norm, fisher_exact

path = sys.argv[1] if len(sys.argv) > 1 else "data/outreach_log_SAMPLE.csv"
df = pd.read_csv(path)
df["replied"] = (df["replied"].str.upper() == "Y").astype(int)
df["meeting"] = (df["meeting_booked"].str.upper() == "Y").astype(int)

s = df.groupby("variant").agg(sent=("replied", "size"), replies=("replied", "sum"), meetings=("meeting", "sum"))
s["reply_rate_pct"] = (s.replies / s.sent * 100).round(1)
s["meeting_rate_pct"] = (s.meetings / s.sent * 100).round(1)
print(s, "\n")

a, b = s.loc["A"], s.loc["B"]
p1, p2 = a.replies / a.sent, b.replies / b.sent
pool = (a.replies + b.replies) / (a.sent + b.sent)
z = (p1 - p2) / np.sqrt(pool * (1 - pool) * (1 / a.sent + 1 / b.sent))
p_z = 2 * (1 - norm.cdf(abs(z)))
_, p_f = fisher_exact([[a.replies, a.sent - a.replies], [b.replies, b.sent - b.replies]])
print(f"Reply rate A={p1:.1%}  B={p2:.1%}  diff={(p1 - p2) * 100:+.1f} pts")
print(f"two-proportion z={z:.2f}  p={p_z:.3f}   |  Fisher exact p={p_f:.3f}")
print("Significant at 5%" if p_f < 0.05 else "Not significant at 5% - collect more sends")

by = df.groupby(["channel", "variant"]).replied.agg(["size", "mean"]).round(3)
print("\nBy channel x variant:\n", by)
