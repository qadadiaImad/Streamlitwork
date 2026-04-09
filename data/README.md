# Data Folder

This folder contains sample yield curve data for testing the Stale Detection module.

## Configuration

The data folder path is configurable in `generate_sample_data.py`:

```python
DATA_FOLDER = Path("data")  # Change this to use a different folder
```

## Generated Files

### `sample_curves.csv`
- **Format**: Consolidated CSV (single file mode)
- **Structure**: First column = date (DDMMYYYY), remaining columns = horizons (1M, 3M, 6M, 1Y, 2Y, 3Y, 5Y, 7Y, 10Y, 15Y, 20Y, 30Y)
- **Data**: 520 business days (~2 years) starting from 03/01/2022
- **Stale zones injected**:
  - **Day 100-110**: 5Y horizon frozen (10-day streak)
  - **Day 200-203**: Entire curve frozen (3-day streak)
  - **Day 350-365**: Short end (1M, 3M) frozen (15-day streak)
  - **Sporadic**: Single-day stales at days 50, 150, 250, 400, 450

## How to Test

1. **Start the dashboard**: `streamlit run app.py`
2. **Navigate to**: Yield Curve — Stale Detection page
3. **Upload**: `data/sample_curves.csv` in "Single consolidated file" mode
4. **Adjust detection settings**:
   - Tolerance: 0.0 (exact match)
   - Lookback: 1 day
   - Streak threshold: 3 days
5. **Expected results**:
   - ~10-15% stale points detected
   - 3D surface shows red zones at injected stale periods
   - Heatmap highlights the frozen zones
   - Severe stale report shows streaks ≥ 3 days

## Regenerate Data

To create fresh sample data:
```bash
python generate_sample_data.py
```
