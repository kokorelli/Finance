import yfinance as yf
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression
import os

def estimate_missing_data(primary_ticker: str, secondary_ticker1: str, secondary_ticker2: str, startDate: str, endDate: str):
    # Download historical data for all three stocks
    data_primary = yf.download(primary_ticker, start=startDate, end=endDate, interval='1mo')['Adj Close']
    data_secondary1 = yf.download(secondary_ticker1, start=startDate, end=endDate, interval='1mo')['Adj Close']
    data_secondary2 = yf.download(secondary_ticker2, start=startDate, end=endDate, interval='1mo')['Adj Close']
    
    # Calculate the capital gains yield for each stock
    cg_yield_primary = data_primary.pct_change().dropna()
    cg_yield_secondary1 = data_secondary1.pct_change().dropna()
    cg_yield_secondary2 = data_secondary2.pct_change().dropna()

    # Combine data into a single DataFrame and drop NaN values
    df = pd.DataFrame({
        primary_ticker: cg_yield_primary,
        secondary_ticker1: cg_yield_secondary1,
        secondary_ticker2: cg_yield_secondary2
    }).dropna()

    # Separate the independent and dependent variables
    X = df[[secondary_ticker1, secondary_ticker2]]
    y = df[primary_ticker]

    # Perform linear regression
    model = LinearRegression()
    model.fit(X, y)

    # Extract the R-squared and coefficients
    r_squared = model.score(X, y)
    intercept = model.intercept_
    coef_1, coef_2 = model.coef_

    # Calculate adjusted R-squared
    n = len(y)  # Number of observations
    k = X.shape[1]  # Number of predictors
    adj_r_squared = 1 - (1 - r_squared) * (n - 1) / (n - k - 1)

    # Print regression results
    print(f"Adjusted R-squared: {adj_r_squared}")
    print(f"Intercept: {intercept}")
    print(f"Coefficient for {secondary_ticker1}: {coef_1}")
    print(f"Coefficient for {secondary_ticker2}: {coef_2}")

    # Predict missing values for primary stock
    X_full = pd.DataFrame({
        secondary_ticker1: cg_yield_secondary1,
        secondary_ticker2: cg_yield_secondary2
    })
    predicted_yield = model.predict(X_full)

    # Fill missing values in primary stock's capital gains yield using regression
    estimated_yield = cg_yield_primary.combine_first(pd.Series(predicted_yield, index=X_full.index))
    
    # Prepare results for Excel
    result_df = pd.DataFrame({
        'Date': estimated_yield.index.strftime('%Y-%m-%d'),
        f'{primary_ticker} Capital Gains Yield': estimated_yield,
        'Adjusted R-squared': [adj_r_squared] * len(estimated_yield),
        'Intercept': [intercept] * len(estimated_yield),
        f'Coefficient for {secondary_ticker1}': [coef_1] * len(estimated_yield),
        f'Coefficient for {secondary_ticker2}': [coef_2] * len(estimated_yield)
    })

    # Write results to a new Excel sheet
    with pd.ExcelWriter('stock-data-final.xlsx', mode='a') as writer:
        result_df.to_excel(writer, sheet_name=f'{primary_ticker} Estimation', index=False)

    print(f"Estimation for {primary_ticker} completed and saved to Excel.")

# Example - Estimate missing data for Visa using Mastercard and American Express
estimate_missing_data("V", "MA", "AXP", '2007-01-01', '2024-08-01')

# Open the Excel file
os.system('start EXCEL.EXE stock-data-final.xlsx')
