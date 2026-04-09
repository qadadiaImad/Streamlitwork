"""
Formula Engine for Interactive Data Transformations
====================================================
Excel-compatible formulas + Custom yield curve formulas
"""
import pandas as pd
import numpy as np
import re
from typing import Dict, Callable, Any, Optional


class FormulaEngine:
    """
    Engine to parse and evaluate Excel-like formulas and custom formulas
    on yield curve dataframes.
    """
    
    def __init__(self, df: pd.DataFrame, h_labels: list):
        """
        Initialize with the surface dataframe and horizon labels.
        
        Args:
            df: Surface dataframe (dates x horizons)
            h_labels: List of horizon labels for column reference
        """
        self.df = df
        self.h_labels = h_labels
        self.date_labels = [d.strftime('%d/%m/%Y') for d in df.index]
        
        # Excel-compatible functions
        self.excel_funcs = {
            'SUM': self._excel_sum,
            'AVERAGE': self._excel_average,
            'MEAN': self._excel_average,
            'MIN': self._excel_min,
            'MAX': self._excel_max,
            'STD': self._excel_std,
            'STDEV': self._excel_std,
            'VAR': self._excel_var,
            'COUNT': self._excel_count,
            'ABS': self._excel_abs,
            'IF': self._excel_if,
            'ROUND': self._excel_round,
            'POWER': self._excel_power,
            'SQRT': self._excel_sqrt,
            'LN': self._excel_ln,
            'LOG': self._excel_log,
            'EXP': self._excel_exp,
        }
        
        # Custom yield curve functions
        self.custom_funcs = {
            'SPREAD': self._custom_spread,
            'BUTTERFLY': self._custom_butterfly,
            'ROLLDOWN': self._custom_rolldown,
            'CARRY': self._custom_carry,
            'STEEPNESS': self._custom_steepness,
            'CURVATURE': self._custom_curvature,
            'DELTA_Y': self._custom_delta_y,
            'ZSCORE': self._custom_zscore,
            'NORMALIZE': self._custom_normalize,
            'CHANGE': self._custom_change,
            'PCT_CHANGE': self._custom_pct_change,
        }
        
        # Combine all functions
        self.all_funcs = {**self.excel_funcs, **self.custom_funcs}
    
    # ============ Excel Functions ============
    
    def _excel_sum(self, *args):
        """SUM(range) - Sum of values"""
        values = self._get_values(args)
        return np.nansum(values) if len(values) > 0 else np.nan
    
    def _excel_average(self, *args):
        """AVERAGE(range) - Arithmetic mean"""
        values = self._get_values(args)
        return np.nanmean(values) if len(values) > 0 else np.nan
    
    def _excel_min(self, *args):
        """MIN(range) - Minimum value"""
        values = self._get_values(args)
        return np.nanmin(values) if len(values) > 0 else np.nan
    
    def _excel_max(self, *args):
        """MAX(range) - Maximum value"""
        values = self._get_values(args)
        return np.nanmax(values) if len(values) > 0 else np.nan
    
    def _excel_std(self, *args):
        """STD(range) - Standard deviation"""
        values = self._get_values(args)
        return np.nanstd(values) if len(values) > 0 else np.nan
    
    def _excel_var(self, *args):
        """VAR(range) - Variance"""
        values = self._get_values(args)
        return np.nanvar(values) if len(values) > 0 else np.nan
    
    def _excel_count(self, *args):
        """COUNT(range) - Count of numeric values"""
        values = self._get_values(args)
        return len([v for v in values if not np.isnan(v)])
    
    def _excel_abs(self, value):
        """ABS(value) - Absolute value"""
        return abs(float(value))
    
    def _excel_if(self, condition, true_val, false_val):
        """IF(condition, true_val, false_val) - Conditional"""
        return true_val if condition else false_val
    
    def _excel_round(self, value, decimals=0):
        """ROUND(value, decimals) - Round to N decimal places"""
        return round(float(value), int(decimals))
    
    def _excel_power(self, base, exponent):
        """POWER(base, exponent) - Exponentiation"""
        return float(base) ** float(exponent)
    
    def _excel_sqrt(self, value):
        """SQRT(value) - Square root"""
        return np.sqrt(float(value))
    
    def _excel_ln(self, value):
        """LN(value) - Natural logarithm"""
        return np.log(float(value))
    
    def _excel_log(self, value, base=10):
        """LOG(value, base) - Logarithm with base"""
        return np.log(float(value)) / np.log(float(base))
    
    def _excel_exp(self, value):
        """EXP(value) - Exponential (e^x)"""
        return np.exp(float(value))
    
    # ============ Custom Yield Curve Functions ============
    
    def _custom_spread(self, leg1_idx, leg2_idx):
        """
        SPREAD(leg1, leg2) - Calculate spread between two tenors
        Returns time series of (leg2 - leg1)
        """
        try:
            i1, i2 = int(leg1_idx), int(leg2_idx)
            return self.df.iloc[:, i2] - self.df.iloc[:, i1]
        except:
            return pd.Series(np.nan, index=self.df.index)
    
    def _custom_butterfly(self, short_idx, mid_idx, long_idx):
        """
        BUTTERFLY(short, mid, long) - Butterfly spread
        Formula: 2*mid - short - long
        """
        try:
            s, m, l = int(short_idx), int(mid_idx), int(long_idx)
            return 2 * self.df.iloc[:, m] - self.df.iloc[:, s] - self.df.iloc[:, l]
        except:
            return pd.Series(np.nan, index=self.df.index)
    
    def _custom_rolldown(self, tenor_idx, days_forward=30):
        """
        ROLLDOWN(tenor, days) - Expected rolldown return
        Approximation: rate at shorter tenor - current rate
        """
        try:
            idx = int(tenor_idx)
            col = self.df.columns[idx]
            # Find next shorter tenor
            shorter_cols = [c for c in self.df.columns if c < col]
            if shorter_cols:
                shorter_col = max(shorter_cols)
                return self.df[shorter_col] - self.df.iloc[:, idx]
            return pd.Series(np.nan, index=self.df.index)
        except:
            return pd.Series(np.nan, index=self.df.index)
    
    def _custom_carry(self, tenor_idx, funding_idx=0):
        """
        CARRY(tenor, funding) - Carry return
        Formula: rate_tenor - rate_funding
        """
        try:
            t, f = int(tenor_idx), int(funding_idx)
            return self.df.iloc[:, t] - self.df.iloc[:, f]
        except:
            return pd.Series(np.nan, index=self.df.index)
    
    def _custom_steepness(self, short_idx, long_idx):
        """
        STEEPNESS(short, long) - Curve steepness
        Returns time series of (long - short)
        """
        return self._custom_spread(short_idx, long_idx)
    
    def _custom_curvature(self, short_idx, mid_idx, long_idx):
        """
        CURVATURE(short, mid, long) - Curve curvature
        Same as butterfly but normalized
        """
        return self._custom_butterfly(short_idx, mid_idx, long_idx)
    
    def _custom_delta_y(self, tenor_idx, lookback=1):
        """
        DELTA_Y(tenor, lookback) - Change in rate over N days
        """
        try:
            idx = int(tenor_idx)
            lb = int(lookback)
            return self.df.iloc[:, idx].diff(lb)
        except:
            return pd.Series(np.nan, index=self.df.index)
    
    def _custom_zscore(self, tenor_idx, window=30):
        """
        ZSCORE(tenor, window) - Z-score normalization over rolling window
        """
        try:
            idx = int(tenor_idx)
            w = int(window)
            series = self.df.iloc[:, idx]
            rolling_mean = series.rolling(window=w, min_periods=1).mean()
            rolling_std = series.rolling(window=w, min_periods=1).std()
            return (series - rolling_mean) / rolling_std.replace(0, np.nan)
        except:
            return pd.Series(np.nan, index=self.df.index)
    
    def _custom_normalize(self, tenor_idx):
        """
        NORMALIZE(tenor) - Min-max normalize to [0, 1]
        """
        try:
            idx = int(tenor_idx)
            series = self.df.iloc[:, idx]
            min_val, max_val = series.min(), series.max()
            if max_val > min_val:
                return (series - min_val) / (max_val - min_val)
            return pd.Series(0.5, index=self.df.index)
        except:
            return pd.Series(np.nan, index=self.df.index)
    
    def _custom_change(self, tenor_idx, periods=1):
        """
        CHANGE(tenor, periods) - Absolute change
        """
        return self._custom_delta_y(tenor_idx, periods)
    
    def _custom_pct_change(self, tenor_idx, periods=1):
        """
        PCT_CHANGE(tenor, periods) - Percentage change
        """
        try:
            idx = int(tenor_idx)
            p = int(periods)
            return self.df.iloc[:, idx].pct_change(p) * 100
        except:
            return pd.Series(np.nan, index=self.df.index)
    
    # ============ Helper Methods ============
    
    def _get_values(self, args):
        """Extract numeric values from arguments"""
        values = []
        for arg in args:
            if isinstance(arg, (list, tuple, np.ndarray, pd.Series)):
                values.extend([float(v) for v in arg if not np.isnan(v)])
            else:
                try:
                    values.append(float(arg))
                except:
                    pass
        return values
    
    def _resolve_reference(self, ref: str) -> Any:
        """
        Resolve cell/column references like '1M', '5Y', 'A1', 'B2:B10'
        
        Supports:
        - Horizon labels: '1M', '3M', '1Y', '10Y'
        - Column indices: 0, 1, 2
        - Date strings: '03/01/2022'
        """
        ref = str(ref).strip()
        
        # Try horizon label match
        if ref in self.h_labels:
            idx = self.h_labels.index(ref)
            return self.df.iloc[:, idx]
        
        # Try numeric column index
        try:
            idx = int(ref)
            if 0 <= idx < len(self.h_labels):
                return self.df.iloc[:, idx]
        except:
            pass
        
        # Try date reference
        if ref in self.date_labels:
            idx = self.date_labels.index(ref)
            return self.df.iloc[idx, :]
        
        # Try to parse as range (e.g., "0:3" or "1M:10Y")
        if ':' in ref:
            parts = ref.split(':')
            if len(parts) == 2:
                start, end = parts[0], parts[1]
                # Try numeric range
                try:
                    start_idx, end_idx = int(start), int(end)
                    return self.df.iloc[:, start_idx:end_idx+1]
                except:
                    # Try horizon label range
                    if start in self.h_labels and end in self.h_labels:
                        start_idx = self.h_labels.index(start)
                        end_idx = self.h_labels.index(end)
                        return self.df.iloc[:, start_idx:end_idx+1]
        
        # Return as literal value
        try:
            return float(ref)
        except:
            return ref
    
    # ============ Public API ============
    
    def parse_formula(self, formula: str) -> tuple:
        """
        Parse a formula string and return (function_name, args)
        
        Supports:
        - Excel style: =FUNCTION(arg1, arg2, ...)
        - Direct: FUNCTION(arg1, arg2, ...)
        - Operations: A + B, A - B, A * B, A / B
        """
        formula = formula.strip()
        
        # Remove leading = if present
        if formula.startswith('='):
            formula = formula[1:]
        
        # Check for simple operations (A + B, A - B, etc.)
        for op in ['+', '-', '*', '/']:
            if op in formula and '(' not in formula:
                parts = formula.split(op)
                if len(parts) == 2:
                    left = self._resolve_reference(parts[0].strip())
                    right = self._resolve_reference(parts[1].strip())
                    return ('OP', op, left, right)
        
        # Parse function call
        func_match = re.match(r'^(\w+)\s*\((.*)\)$', formula, re.IGNORECASE)
        if func_match:
            func_name = func_match.group(1).upper()
            args_str = func_match.group(2)
            
            # Parse arguments (handle nested parentheses)
            args = self._parse_args(args_str)
            
            return (func_name, args)
        
        # Single reference
        return ('REF', [self._resolve_reference(formula)])
    
    def _parse_args(self, args_str: str) -> list:
        """Parse comma-separated arguments, respecting nested parentheses"""
        args = []
        current = ''
        depth = 0
        
        for char in args_str:
            if char == '(':
                depth += 1
                current += char
            elif char == ')':
                depth -= 1
                current += char
            elif char == ',' and depth == 0:
                args.append(current.strip())
                current = ''
            else:
                current += char
        
        if current.strip():
            args.append(current.strip())
        
        # Resolve each argument
        resolved = []
        for arg in args:
            # Check if arg is a nested formula
            if '(' in arg:
                nested_result = self.evaluate(arg)
                resolved.append(nested_result)
            else:
                resolved.append(self._resolve_reference(arg))
        
        return resolved
    
    def evaluate(self, formula: str) -> Any:
        """
        Evaluate a formula and return the result
        """
        try:
            parsed = self.parse_formula(formula)
            
            # Handle simple reference
            if parsed[0] == 'REF':
                return parsed[1][0]
            
            # Handle operation
            if parsed[0] == 'OP':
                _, op, left, right = parsed
                if isinstance(left, pd.Series) and isinstance(right, pd.Series):
                    if op == '+': return left + right
                    if op == '-': return left - right
                    if op == '*': return left * right
                    if op == '/': return left / right
                elif isinstance(left, pd.Series):
                    if op == '+': return left + float(right)
                    if op == '-': return left - float(right)
                    if op == '*': return left * float(right)
                    if op == '/': return left / float(right)
                elif isinstance(right, pd.Series):
                    if op == '+': return float(left) + right
                    if op == '-': return float(left) - right
                    if op == '*': return float(left) * right
                    if op == '/': return float(left) / right
                else:
                    if op == '+': return float(left) + float(right)
                    if op == '-': return float(left) - float(right)
                    if op == '*': return float(left) * float(right)
                    if op == '/': return float(left) / float(right)
                return None
            
            # Handle function call
            func_name, args = parsed
            
            if func_name in self.all_funcs:
                result = self.all_funcs[func_name](*args)
                
                # If result is a Series, ensure it has proper index
                if isinstance(result, pd.Series):
                    result.index = self.df.index
                
                return result
            else:
                return f"Error: Unknown function '{func_name}'"
        
        except Exception as e:
            return f"Error: {str(e)}"
    
    def get_available_functions(self) -> Dict[str, str]:
        """Get documentation for all available functions"""
        docs = {
            # Excel functions
            'SUM(range)': 'Sum of values',
            'AVERAGE(range)': 'Arithmetic mean',
            'MIN(range)': 'Minimum value',
            'MAX(range)': 'Maximum value',
            'STD(range)': 'Standard deviation',
            'COUNT(range)': 'Count of values',
            'ABS(value)': 'Absolute value',
            'IF(cond, true, false)': 'Conditional value',
            'ROUND(value, n)': 'Round to N decimals',
            'POWER(base, exp)': 'Exponentiation',
            'SQRT(value)': 'Square root',
            'LN(value)': 'Natural log',
            'LOG(value, base)': 'Logarithm',
            'EXP(value)': 'Exponential (e^x)',
            
            # Custom functions
            'SPREAD(leg1, leg2)': 'Spread between two tenors (leg2 - leg1)',
            'BUTTERFLY(s, m, l)': 'Butterfly: 2*mid - short - long',
            'ROLLDOWN(tenor, days)': 'Expected rolldown return',
            'CARRY(tenor, funding)': 'Carry return (tenor - funding)',
            'STEEPNESS(s, l)': 'Curve steepness (long - short)',
            'CURVATURE(s, m, l)': 'Curve curvature',
            'DELTA_Y(tenor, lookback)': 'Rate change over N days',
            'ZSCORE(tenor, window)': 'Z-score over rolling window',
            'NORMALIZE(tenor)': 'Min-max normalize to [0,1]',
            'CHANGE(tenor, periods)': 'Absolute change',
            'PCT_CHANGE(tenor, periods)': 'Percentage change',
        }
        return docs


def apply_formula_to_surface(surface_df: pd.DataFrame, h_labels: list, 
                              formula: str, new_column_name: str = None) -> pd.DataFrame:
    """
    Apply a formula to the surface dataframe and return a new column
    
    Args:
        surface_df: Original surface dataframe
        h_labels: Horizon labels
        formula: Formula string to evaluate
        new_column_name: Name for the new column (optional)
    
    Returns:
        DataFrame with new column added
    """
    engine = FormulaEngine(surface_df, h_labels)
    result = engine.evaluate(formula)
    
    if isinstance(result, pd.Series):
        df = surface_df.copy()
        col_name = new_column_name or f"Calc_{len(df.columns)}"
        df[col_name] = result
        return df
    
    return surface_df
