"""
Generate sample yield curve data for testing the stale detection dashboard.
Run: python generate_sample_data.py
Creates: sample_curves.csv with 2 years of daily data + intentional stale zones.
"""
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pathlib import Path

np.random.seed(42)

# Config
DATA_FOLDER = Path("data")  # Configurable data folder path
DATA_FOLDER.mkdir(exist_ok=True)  # Create folder if it doesn't exist

start_date = datetime(2022, 1, 3)
n_days = 520  # ~2 years of business days
horizons = ['1M', '3M', '6M', '1Y', '2Y', '3Y', '5Y', '7Y', '10Y', '15Y', '20Y', '30Y']
horizon_years = [1/12, 3/12, 6/12, 1, 2, 3, 5, 7, 10, 15, 20, 30]

# Base curve (Nelson-Siegel-like)
def base_curve(t, beta0=0.04, beta1=-0.02, beta2=0.01, tau=2.0):
    factor = (1 - np.exp(-t/tau)) / (t/tau + 1e-10)
    return beta0 + beta1 * factor + beta2 * (factor - np.exp(-t/tau))

dates = []
dt = start_date
while len(dates) < n_days:
    if dt.weekday() < 5:  # Business days only
        dates.append(dt)
    dt += timedelta(days=1)

data = {}
for i, date in enumerate(dates):
    # Evolving curve with random walk
    shift = 0.001 * np.random.randn()
    slope_shift = 0.0005 * np.random.randn()
    rates = []
    for j, t in enumerate(horizon_years):
        r = base_curve(t, beta0=0.04 + 0.01*np.sin(2*np.pi*i/260))
        r += shift + slope_shift * t
        r += 0.0003 * np.random.randn()  # Noise
        rates.append(round(r, 6))
    data[date.strftime('%d%m%Y')] = rates

df = pd.DataFrame(data, index=horizons).T
df.index.name = 'date'

# Inject stale zones intentionally
# Zone 1: 5Y frozen for 10 days around day 100
for i in range(100, 110):
    df.iloc[i, horizons.index('5Y')] = df.iloc[99, horizons.index('5Y')]

# Zone 2: Entire curve frozen for 3 days around day 200
for i in range(200, 203):
    df.iloc[i] = df.iloc[199]

# Zone 3: Short end (1M, 3M) frozen for 15 days around day 350
for i in range(350, 365):
    df.iloc[i, horizons.index('1M')] = df.iloc[349, horizons.index('1M')]
    df.iloc[i, horizons.index('3M')] = df.iloc[349, horizons.index('3M')]

# Zone 4: Sporadic single-day stales
for d in [50, 150, 250, 400, 450]:
    df.iloc[d] = df.iloc[d-1]

output_file = DATA_FOLDER / 'sample_curves.csv'
df.to_csv(output_file)
print(f"✅ Generated {output_file}: {len(df)} dates × {len(horizons)} horizons")
print(f"   Stale zones injected at days 100-110 (5Y), 200-203 (all), 350-365 (1M/3M), and sporadic singles.")
print(f"\n   Upload this file in 'Single consolidated file' mode in the Yield Curve dashboard.")
