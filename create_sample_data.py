#!/usr/bin/env python3
"""
Script to create sample CSV data for testing the Company Email Generator
Run with: python create_sample_data.py
"""

import pandas as pd
import random
from datetime import datetime, timedelta

def create_sample_data():
    """Create a comprehensive sample dataset with companies from different exchanges"""
    
    # Sample data for different exchanges and sectors
    companies_data = [
        # NASDAQ Tech Companies
        {"Company": "Meta Platforms Inc", "Ticker": "META", "Exchange": "NASDAQ", "Sector": "Technology", "Industry": "Social Media", "Country": "United States", "City": "Menlo Park CA"},
        {"Company": "Netflix Inc", "Ticker": "NFLX", "Exchange": "NASDAQ", "Sector": "Communication Services", "Industry": "Entertainment", "Country": "United States", "City": "Los Gatos CA"},
        {"Company": "Adobe Inc", "Ticker": "ADBE", "Exchange": "NASDAQ", "Sector": "Technology", "Industry": "Software", "Country": "United States", "City": "San Jose CA"},
        
        # NYSE Companies
        {"Company": "Coca-Cola Company", "Ticker": "KO", "Exchange": "NYSE", "Sector": "Consumer Staples", "Industry": "Beverages", "Country": "United States", "City": "Atlanta GA"},
        {"Company": "Procter & Gamble Co", "Ticker": "PG", "Exchange": "NYSE", "Sector": "Consumer Staples", "Industry": "Personal Products", "Country": "United States", "City": "Cincinnati OH"},
        {"Company": "Visa Inc", "Ticker": "V", "Exchange": "NYSE", "Sector": "Financials", "Industry": "Payment Processing", "Country": "United States", "City": "San Francisco CA"},
        
        # TSX Canadian Companies
        {"Company": "Canadian Pacific Railway", "Ticker": "CP", "Exchange": "TSX", "Sector": "Industrials", "Industry": "Transportation", "Country": "Canada", "City": "Calgary AB"},
        {"Company": "Nutrien Ltd", "Ticker": "NTR", "Exchange": "TSX", "Sector": "Materials", "Industry": "Fertilizers", "Country": "Canada", "City": "Saskatoon SK"},
        {"Company": "Canadian Tire Corporation", "Ticker": "CTC", "Exchange": "TSX", "Sector": "Consumer Discretionary", "Industry": "Retail", "Country": "Canada", "City": "Toronto ON"},
        
        # LSE UK Companies
        {"Company": "Shell plc", "Ticker": "SHEL", "Exchange": "LSE", "Sector": "Energy", "Industry": "Oil & Gas", "Country": "United Kingdom", "City": "London"},
        {"Company": "HSBC Holdings plc", "Ticker": "HSBA", "Exchange": "LSE", "Sector": "Financials", "Industry": "Banking", "Country": "United Kingdom", "City": "London"},
        {"Company": "British American Tobacco", "Ticker": "BATS", "Exchange": "LSE", "Sector": "Consumer Staples", "Industry": "Tobacco", "Country": "United Kingdom", "City": "London"},
        
        # ASX Australian Companies
        {"Company": "Telstra Corporation", "Ticker": "TLS", "Exchange": "ASX", "Sector": "Communication Services", "Industry": "Telecommunications", "Country": "Australia", "City": "Melbourne VIC"},
        {"Company": "Wesfarmers Limited", "Ticker": "WES", "Exchange": "ASX", "Sector": "Consumer Discretionary", "Industry": "Retail", "Country": "Australia", "City": "Perth WA"},
        {"Company": "Macquarie Group", "Ticker": "MQG", "Exchange": "ASX", "Sector": "Financials", "Industry": "Investment Banking", "Country": "Australia", "City": "Sydney NSW"},
        
        # TSX-V Smaller Companies
        {"Company": "Lithium Americas Corp", "Ticker": "LAC", "Exchange": "TSX-V", "Sector": "Materials", "Industry": "Lithium Mining", "Country": "Canada", "City": "Vancouver BC"},
        {"Company": "Green Thumb Industries", "Ticker": "GTII", "Exchange": "TSX-V", "Sector": "Healthcare", "Industry": "Cannabis", "Country": "Canada", "City": "Toronto ON"},
    ]
    
    # Generate additional fields for each company
    for company in companies_data:
        # Market Cap (random but sector-appropriate)
        if company["Sector"] == "Technology":
            market_cap = random.randint(50000000000, 3000000000000)  # 50B - 3T
        elif company["Sector"] == "Energy":
            market_cap = random.randint(20000000000, 500000000000)   # 20B - 500B
        elif company["Exchange"] == "TSX-V":
            market_cap = random.randint(100000000, 5000000000)       # 100M - 5B
        else:
            market_cap = random.randint(5000000000, 200000000000)    # 5B - 200B
        
        company["Market_Cap_USD"] = market_cap
        company["Current_Stock_Price"] = round(random.uniform(10, 500), 2)
        
        # Executive information
        company["CEO_Name"] = f"John Smith"  # Placeholder
        company["CEO_Email"] = f"ceo@{company['Ticker'].lower()}.com"
        company["CFO_Name"] = f"Jane Doe"    # Placeholder
        company["CFO_Email"] = f"cfo@{company['Ticker'].lower()}.com"
        company["IR_Contact_Name"] = f"IR Team"
        company["IR_Contact_Email"] = f"investor.relations@{company['Ticker'].lower()}.com"
        
        # Australian subsidiary (higher chance for US/UK companies)
        if company["Country"] in ["United States", "United Kingdom"]:
            company["Has_Australian_Subsidiary"] = random.choice(["Yes", "Yes", "No"])
        elif company["Country"] == "Australia":
            company["Has_Australian_Subsidiary"] = "No"  # They ARE Australian
        else:
            company["Has_Australian_Subsidiary"] = random.choice(["Yes", "No", "No"])
        
        if company["Has_Australian_Subsidiary"] == "Yes":
            company["Australian_Subsidiary_Name"] = f"{company['Company'].split()[0]} Australia"
            company["Australian_Office_Location"] = random.choice(["Sydney NSW", "Melbourne VIC", "Perth WA", "Brisbane QLD"])
        else:
            company["Australian_Subsidiary_Name"] = ""
            company["Australian_Office_Location"] = ""
        
        # Capital raising data (last 3 years)
        company["Capital_Raised_2023"] = random.randint(0, 5000000000) if random.random() > 0.5 else 0
        company["Capital_Raised_2022"] = random.randint(0, 8000000000) if random.random() > 0.4 else 0
        company["Capital_Raised_2021"] = random.randint(0, 12000000000) if random.random() > 0.3 else 0
        company["Capital_Raised_5_Year_Total"] = company["Capital_Raised_2023"] + company["Capital_Raised_2022"] + company["Capital_Raised_2021"] + random.randint(0, 15000000000)
        
        # Listing date (random historical date)
        start_date = datetime(1980, 1, 1)
        end_date = datetime(2020, 12, 31)
        random_date = start_date + timedelta(days=random.randint(0, (end_date - start_date).days))
        company["Primary_Listing_Date"] = random_date.strftime("%Y-%m-%d")
        
        # Additional financial metrics
        company["Website"] = f"https://www.{company['Ticker'].lower()}.com"
        company["Employee_Count"] = random.randint(1000, 300000)
        company["Revenue_USD_Latest"] = random.randint(int(market_cap * 0.1), int(market_cap * 0.8))
        company["Net_Income_USD_Latest"] = random.randint(int(company["Revenue_USD_Latest"] * 0.05), int(company["Revenue_USD_Latest"] * 0.25))
        company["Total_Assets_USD"] = random.randint(int(market_cap * 0.8), int(market_cap * 2))
        company["Total_Debt_USD"] = random.randint(int(company["Total_Assets_USD"] * 0.1), int(company["Total_Assets_USD"] * 0.4))
        company["Dividend_Yield_Percent"] = round(random.uniform(0, 8), 1) if random.random() > 0.3 else 0.0
        company["PE_Ratio"] = round(random.uniform(5, 50), 1) if company["Net_Income_USD_Latest"] > 0 else 0.0
        company["Business_Description"] = f"Leading {company['Industry'].lower()} company operating in the {company['Sector'].lower()} sector"
    
    return companies_data

def save_to_csv(data, filename):
    """Save data to CSV file"""
    df = pd.DataFrame(data)
    df.to_csv(filename, index=False)
    print(f"✅ Created {filename} with {len(data)} companies")
    print(f"📊 Exchanges: {', '.join(df['Exchange'].unique())}")
    print(f"🌍 Countries: {', '.join(df['Country'].unique())}")
    print(f"🏢 Sectors: {', '.join(df['Sector'].unique())}")

if __name__ == "__main__":
    print("🏭 Creating sample company database...")
    
    # Create comprehensive sample data
    sample_data = create_sample_data()
    
    # Save to test_data directory
    import os
    os.makedirs("test_data", exist_ok=True)
    
    save_to_csv(sample_data, "test_data/sample_global_companies.csv")
    
    # Create filtered datasets for specific use cases
    
    # US companies only
    us_companies = [c for c in sample_data if c["Country"] == "United States"]
    save_to_csv(us_companies, "test_data/us_companies_only.csv")
    
    # Companies with Australian subsidiaries
    aus_sub_companies = [c for c in sample_data if c["Has_Australian_Subsidiary"] == "Yes"]
    save_to_csv(aus_sub_companies, "test_data/companies_with_aus_subsidiaries.csv")
    
    # Technology sector only
    tech_companies = [c for c in sample_data if c["Sector"] == "Technology"]
    save_to_csv(tech_companies, "test_data/technology_companies.csv")
    
    # Create sample data with real ticker symbols for testing PowerPoint functionality
    data = {
        'Company': ['Apple Inc.', 'Microsoft Corporation', 'Amazon.com Inc.', 'Alphabet Inc.', 'Tesla Inc.', 'NVIDIA Corporation'],
        'Ticker': ['AAPL', 'MSFT', 'AMZN', 'GOOGL', 'TSLA', 'NVDA'], 
        'Exchange': ['NASDAQ', 'NASDAQ', 'NASDAQ', 'NASDAQ', 'NASDAQ', 'NASDAQ'],
        'Country': ['USA', 'USA', 'USA', 'USA', 'USA', 'USA'],
        'Sector': ['Technology', 'Technology', 'Consumer Discretionary', 'Technology', 'Consumer Discretionary', 'Technology'],
        'Industry': ['Consumer Electronics', 'Software', 'E-commerce', 'Internet Services', 'Electric Vehicles', 'Semiconductors'],
        'Market_Cap_USD': [3000000000000, 2800000000000, 1500000000000, 1700000000000, 800000000000, 2200000000000],
        'Has_Australian_Subsidiary': ['Yes', 'Yes', 'Yes', 'Yes', 'Yes', 'No'],
        'CEO_Name': ['Tim Cook', 'Satya Nadella', 'Andy Jassy', 'Sundar Pichai', 'Elon Musk', 'Jensen Huang'],
        'CEO_Email': ['tcook@apple.com', 'satyan@microsoft.com', 'ajassy@amazon.com', 'sundar@google.com', 'elon@tesla.com', 'jhuang@nvidia.com'],
        'IR_Contact_Email': ['investor@apple.com', 'investor@microsoft.com', 'ir@amazon.com', 'investor@alphabet.com', 'ir@tesla.com', 'ir@nvidia.com'],
        'City': ['Cupertino', 'Redmond', 'Seattle', 'Mountain View', 'Austin', 'Santa Clara']
    }

    df = pd.DataFrame(data)
    df.to_csv('test_companies_with_tickers.csv', index=False)
    print("Sample data with real tickers created: test_companies_with_tickers.csv")
    print(f"Created {len(df)} companies for testing PowerPoint functionality")
    print("\nCompanies included:")
    for _, row in df.iterrows():
        print(f"- {row['Company']} ({row['Ticker']}) - {row['Market_Cap_USD']/1e9:.0f}B market cap")
    
    print("\n🎉 Sample data creation complete!")
    print("📁 Files created in test_data/ directory:")
    print("   - sample_global_companies.csv (all companies)")
    print("   - us_companies_only.csv (US companies)")
    print("   - companies_with_aus_subsidiaries.csv (companies with Australian presence)")
    print("   - technology_companies.csv (tech sector only)")
    print("   - test_companies_with_tickers.csv (companies with real tickers)") 