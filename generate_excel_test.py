"""
Generate test Excel files (.xlsb, .xlsm, .xlsx) for testing the Excel Tools module.
Run: python generate_excel_test.py
"""
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime, timedelta

# Configurable data folder
DATA_FOLDER = Path("data")
DATA_FOLDER.mkdir(exist_ok=True)

np.random.seed(42)

# Create sample data
dates = pd.date_range(start='2024-01-01', periods=100, freq='B')  # 100 business days
categories = ['Equity', 'Fixed Income', 'FX', 'Commodities', 'Derivatives']

# Sheet 1: Daily PnL
df_pnl = pd.DataFrame({
    'Date': dates,
    'PnL_USD': np.random.randn(100) * 100000,
    'PnL_EUR': np.random.randn(100) * 85000,
    'PnL_GBP': np.random.randn(100) * 75000,
})
df_pnl['Total_PnL'] = df_pnl['PnL_USD'] + df_pnl['PnL_EUR'] + df_pnl['PnL_GBP']

# Sheet 2: Position data
df_positions = pd.DataFrame({
    'Trade_ID': [f'TRD_{i:05d}' for i in range(1, 51)],
    'Category': np.random.choice(categories, 50),
    'Notional': np.random.uniform(1e6, 50e6, 50).round(2),
    'Maturity': pd.date_range(start='2024-03-01', periods=50, freq='W'),
    'Delta': np.random.randn(50).round(4),
    'Gamma': np.random.randn(50).round(6),
    'Vega': np.random.randn(50).round(2),
})

# Sheet 3: Risk metrics
df_risk = pd.DataFrame({
    'Metric': ['VaR_95', 'VaR_99', 'Expected_Shortfall', 'Max_Drawdown', 'Sharpe_Ratio', 'Beta'],
    'Value': [2.5, 4.2, 5.8, 12.3, 1.45, 0.92],
    'Unit': ['M USD', 'M USD', 'M USD', '%', 'ratio', 'ratio'],
    'Limit': [3.0, 5.0, 7.0, 15.0, 1.0, 1.5],
    'Status': ['OK', 'OK', 'Warning', 'OK', 'OK', 'OK'],
})

# Create .xlsx file (standard)
excel_path = DATA_FOLDER / 'sample_trades.xlsx'
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    df_pnl.to_excel(writer, sheet_name='Daily_PnL', index=False)
    df_positions.to_excel(writer, sheet_name='Positions', index=False)
    df_risk.to_excel(writer, sheet_name='Risk_Metrics', index=False)
print(f"✅ Created {excel_path}")

# Create .xlsb file (binary format - requires pyxlsb engine for reading)
# Note: pandas ExcelWriter doesn't support writing xlsb directly
# We'll create xlsx and note that xlsb requires different approach
# For true xlsb testing, we need to use a different method or manually create

# Create CSV versions for reference
df_pnl.to_csv(DATA_FOLDER / 'sample_pnl.csv', index=False)
df_positions.to_csv(DATA_FOLDER / 'sample_positions.csv', index=False)
df_risk.to_csv(DATA_FOLDER / 'sample_risk.csv', index=False)
print(f"✅ Created CSV versions for reference")

# Try to create actual xlsb using xlwings (requires Excel installed)
try:
    import xlwings as xw
    app = xw.App(visible=False)
    wb = app.books.open(str(excel_path))
    xlsb_path = DATA_FOLDER / 'sample_trades.xlsb'
    wb.save(str(xlsb_path))
    wb.close()
    app.quit()
    print(f"✅ Created {xlsb_path} (binary format)")
except Exception as e:
    print(f"⚠️ Could not create xlsb via xlwings: {e}")
    print("   (Excel may not be installed or accessible)")

# Try to create xlsm (macro-enabled) using xlwings
try:
    import xlwings as xw
    app = xw.App(visible=False)
    wb = app.books.open(str(excel_path))
    # Add a simple macro module
    try:
        module = wb.api.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        module.Name = "TestModule"
        macro_code = '''
Sub HelloWorld()
    MsgBox "Hello from VBA!"
End Sub

Function DoubleValue(x As Double) As Double
    DoubleValue = x * 2
End Function
'''
        module.CodeModule.AddFromString(macro_code)
    except Exception as e:
        print(f"⚠️ Could not add VBA macro: {e}")
    
    xlsm_path = DATA_FOLDER / 'sample_trades.xlsm'
    wb.save(str(xlsm_path))
    wb.close()
    app.quit()
    print(f"✅ Created {xlsm_path} (macro-enabled)")
except Exception as e:
    print(f"⚠️ Could not create xlsm via xlwings: {e}")

print(f"\n📁 All files created in: {DATA_FOLDER.absolute()}")
print(f"\n🧪 Test Excel Tools module by uploading these files:")
print(f"   - sample_trades.xlsx (standard Excel)")
print(f"   - sample_trades.xlsb (binary format - if created)")
print(f"   - sample_trades.xlsm (macro-enabled - if created)")
print(f"   - sample_pnl.csv (CSV alternative)")
