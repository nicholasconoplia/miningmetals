# main.py

import os
from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import requests
from openai import OpenAI
from dotenv import load_dotenv

# 1. Load environment variables from .env
load_dotenv()
NEWSAPI_KEY = os.getenv("NEWSAPI_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# 2. Initialize Flask
app = Flask(__name__)

# 3. Initialize OpenAI client
client = OpenAI(api_key=OPENAI_API_KEY)

def find_column(df, possible_names):
    """
    Helper function to find a column from a list of possible names
    """
    for name in possible_names:
        if name in df.columns:
            return name
    return None


# -----------------------------
# ROUTE #1: HOME / UPLOAD PAGE
# -----------------------------
@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        # A. Check if file part is in request
        if "file" not in request.files:
            return "No file part in the request", 400

        file = request.files["file"]
        if file.filename == "":
            return "No file selected", 400

        # B. Save the file to a temp location
        filepath = os.path.join("uploads", file.filename)
        os.makedirs("uploads", exist_ok=True)
        file.save(filepath)

        # C. Read the file with pandas (support CSV & Excel) - Enhanced with robust parsing
        try:
            if filepath.lower().endswith(".csv"):
                # Try multiple approaches for robust CSV reading
                try:
                    # First attempt: Standard reading
                    df = pd.read_csv(filepath)
                except Exception as e1:
                    try:
                        # Second attempt: Handle malformed CSV with different options
                        df = pd.read_csv(filepath, 
                                       encoding='utf-8',
                                       on_bad_lines='skip')
                    except Exception as e2:
                        try:
                            # Third attempt: More lenient parsing
                            df = pd.read_csv(filepath, 
                                           sep=',',
                                           quotechar='"',
                                           skipinitialspace=True,
                                           on_bad_lines='skip',
                                           encoding='utf-8')
                        except Exception as e3:
                            try:
                                # Fourth attempt: Try different encoding
                                df = pd.read_csv(filepath, 
                                               on_bad_lines='skip',
                                               encoding='latin-1')
                            except Exception as e4:
                                try:
                                    # Fifth attempt: Very lenient parsing for older pandas
                                    df = pd.read_csv(filepath, 
                                                   encoding='utf-8',
                                                   sep=None,
                                                   engine='python')
                                except Exception as e5:
                                    return f"Error reading CSV file. Please check file format. Main error: {str(e1)[:150]}", 400
            else:
                df = pd.read_excel(filepath)  # requires openpyxl installed
        except Exception as e:
            return f"Error reading file: {e}", 400

        # D. Intelligent column detection - find the main company identifier column
        company_column = None
        possible_company_names = ['Company', 'company', 'Company Name', 'company_name', 
                                'CompanyName', 'Name', 'name', 'Symbol', 'Ticker', 
                                'Company_Name', 'COMPANY', 'COMPANY_NAME', 'Issuer',
                                'Company Legal Name', 'Legal Name']
        
        for col_name in possible_company_names:
            if col_name in df.columns:
                company_column = col_name
                break
        
        if company_column is None:
            # If no standard company column found, show available columns
            available_cols = ', '.join(df.columns.tolist())
            return f"Could not find a company name column. Available columns: {available_cols}. Please ensure your file has a column named 'Company', 'Name', or similar.", 400

        # E. Enhanced column standardization - map various column names to standard ones
        column_mappings = {
            # Company name (already handled above)
            'Company': company_column,
            
            # Exchange mappings
            'Exchange': find_column(df, ['Exchange', 'Stock_Exchange', 'Listed_Exchange', 'Market']),
            
            # Sector mappings
            'Sector': find_column(df, ['Sector', 'Primary Industry Sector', 'Industry_Sector', 'Business_Sector', 'Sector_Name']),
            
            # Industry mappings  
            'Industry': find_column(df, ['Industry', 'Primary Industry Group', 'Industry_Group', 'Sub_Industry', 'Industry_Name', 'Business_Industry']),
            
            # Country mappings
            'Country': find_column(df, ['Country', 'HQ Global Country/Territory', 'HQ_Country', 'Headquarters_Country', 'Location_Country']),
            
            # City mappings
            'City': find_column(df, ['City', 'HQ City', 'HQ_City', 'Headquarters_City', 'Location', 'Office_Location']),
            
            # Market cap mappings
            'Market_Cap_USD': find_column(df, ['Market_Cap_USD', 'Market Cap', 'Market_Cap', 'MarketCap', 'Market_Capitalization']),
            
            # Australian subsidiary mappings
            'Has_Australian_Subsidiary': find_column(df, ['Has_Australian_Subsidiary', 'Australian_Subsidiary', 'AUS_Subsidiary', 'Australia_Operations'])
        }
        
        # Apply column mappings (rename columns to standard names)
        rename_dict = {}
        for standard_name, source_column in column_mappings.items():
            if source_column and source_column in df.columns and source_column != standard_name:
                rename_dict[source_column] = standard_name
        
        if rename_dict:
            df = df.rename(columns=rename_dict)
            print(f"Mapped columns: {rename_dict}")
            
            # Update company_column if it was renamed
            if company_column in rename_dict:
                company_column = rename_dict[company_column]
        
        # F. Clean and validate the data
        # Remove rows where company name is missing
        df = df.dropna(subset=[company_column])
        df = df[df[company_column].astype(str).str.strip() != '']
        
        if len(df) == 0:
            return "No valid company data found in the file.", 400
        
        # Standardize the company column name to 'Company' for consistency
        if company_column != 'Company':
            df = df.rename(columns={company_column: 'Company'})
        
        # G. Show user what columns were detected
        detected_columns = df.columns.tolist()
        print(f"Successfully loaded {len(df)} companies with columns: {detected_columns}")
        
        # G. Store `df` in session‐like place with cleaned data
        df.to_csv("uploads/last_upload.csv", index=False)

        # F. Redirect to the "select companies" page
        return redirect(url_for("select_companies"))

    # If GET, just render the upload form
    return render_template("upload.html")


# ----------------------------------
# ROUTE #2: SELECT COMPANIES TO PROCESS
# ----------------------------------
@app.route("/select", methods=["GET", "POST"])
def select_companies():
    # A. Load the stored DataFrame
    try:
        df = pd.read_csv("uploads/last_upload.csv")
    except FileNotFoundError:
        return redirect(url_for("upload_file"))

    # B. Apply filters if provided - Enhanced to handle missing columns
    filtered_df = df.copy()
    
    # Helper function to safely get column values
    def safe_filter(dataframe, column_name, filter_value):
        if column_name in dataframe.columns and filter_value:
            return dataframe[dataframe[column_name] == filter_value]
        return dataframe
    
    def safe_filter_boolean(dataframe, column_name, filter_value):
        if column_name in dataframe.columns and filter_value:
            if filter_value == 'Yes':
                return dataframe[dataframe[column_name] == 'Yes']
            elif filter_value == 'No':
                return dataframe[dataframe[column_name] == 'No']
        return dataframe
    
    def safe_filter_market_cap(dataframe, filter_value):
        # Find market cap column
        market_cap_cols = ['Market_Cap_USD', 'Market_Cap', 'MarketCap', 'Market_Capitalization']
        market_cap_col = None
        for col in market_cap_cols:
            if col in dataframe.columns:
                market_cap_col = col
                break
        
        if not market_cap_col or not filter_value:
            return dataframe
        
        try:
            # Convert market cap to numeric values (handle various formats)
            def parse_market_cap(value):
                if pd.isna(value):
                    return 0
                value_str = str(value).replace(',', '').replace('$', '').strip()
                if 'B' in value_str.upper():
                    return float(value_str.upper().replace('B', '')) * 1000000000
                elif 'M' in value_str.upper():
                    return float(value_str.upper().replace('M', '')) * 1000000
                else:
                    try:
                        return float(value_str)
                    except:
                        return 0
            
            dataframe['_market_cap_numeric'] = dataframe[market_cap_col].apply(parse_market_cap)
            
            if filter_value == 'under_100m':
                return dataframe[dataframe['_market_cap_numeric'] < 100000000]
            elif filter_value == '100m_500m':
                return dataframe[(dataframe['_market_cap_numeric'] >= 100000000) & (dataframe['_market_cap_numeric'] < 500000000)]
            elif filter_value == '500m_1b':
                return dataframe[(dataframe['_market_cap_numeric'] >= 500000000) & (dataframe['_market_cap_numeric'] < 1000000000)]
            elif filter_value == '1b_5b':
                return dataframe[(dataframe['_market_cap_numeric'] >= 1000000000) & (dataframe['_market_cap_numeric'] < 5000000000)]
            elif filter_value == 'over_5b':
                return dataframe[dataframe['_market_cap_numeric'] >= 5000000000]
            
        except Exception as e:
            print(f"Market cap filtering error: {e}")
            return dataframe
        
        return dataframe
    
    def safe_filter_location(dataframe, filter_value):
        # Find location columns
        location_cols = ['City', 'Location', 'Headquarters', 'Australian_Office_Location', 'Office_Location']
        
        if not filter_value:
            return dataframe
        
        location_mask = pd.Series([False] * len(dataframe), index=dataframe.index)
        
        for col in location_cols:
            if col in dataframe.columns:
                # Case-insensitive partial matching
                mask = dataframe[col].astype(str).str.contains(filter_value, case=False, na=False)
                location_mask = location_mask | mask
        
        return dataframe[location_mask] if location_mask.any() else dataframe
    
    # Filter options with safe checking
    filter_exchange = request.args.get('exchange', '')
    filter_country = request.args.get('country', '')
    filter_australian_sub = request.args.get('australian_sub', '')
    filter_sector = request.args.get('sector', '')
    filter_industry = request.args.get('industry', '')  # New industry filter
    filter_market_cap = request.args.get('market_cap_range', '')  # New market cap filter
    filter_location = request.args.get('location', '')  # New location filter
    
    # Apply filters only if columns exist
    filtered_df = safe_filter(filtered_df, 'Exchange', filter_exchange)
    filtered_df = safe_filter(filtered_df, 'Country', filter_country)
    filtered_df = safe_filter_boolean(filtered_df, 'Has_Australian_Subsidiary', filter_australian_sub)
    filtered_df = safe_filter(filtered_df, 'Sector', filter_sector)
    filtered_df = safe_filter(filtered_df, 'Industry', filter_industry)  # Industry filter
    filtered_df = safe_filter_market_cap(filtered_df, filter_market_cap)  # Market cap filter
    filtered_df = safe_filter_location(filtered_df, filter_location)  # Location filter

    companies = filtered_df["Company"].tolist()
    companies_with_details = []
    
    # Get additional details for display - handle missing columns gracefully
    for idx, row in filtered_df.iterrows():
        company_info = {
            'original_index': df.index[df['Company'] == row['Company']].tolist()[0],
            'name': row['Company'],
            'exchange': row.get('Exchange', 'N/A'),
            'country': row.get('Country', 'N/A'),
            'sector': row.get('Sector', 'N/A'),
            'industry': row.get('Industry', 'N/A'),
            'city': row.get('City', row.get('Location', 'N/A')),
            'market_cap': row.get('Market_Cap_USD', row.get('Market_Cap', row.get('MarketCap', 'N/A'))),
            'has_aus_sub': row.get('Has_Australian_Subsidiary', row.get('Australian_Subsidiary', 'N/A'))
        }
        companies_with_details.append(company_info)

    # Get unique values for filter dropdowns - only for existing columns
    exchanges = sorted([x for x in df['Exchange'].unique() if pd.notna(x)]) if 'Exchange' in df.columns else []
    countries = sorted([x for x in df['Country'].unique() if pd.notna(x)]) if 'Country' in df.columns else []
    sectors = sorted([x for x in df['Sector'].unique() if pd.notna(x)]) if 'Sector' in df.columns else []
    industries = sorted([x for x in df['Industry'].unique() if pd.notna(x)]) if 'Industry' in df.columns else []
    
    # Generate dynamic location options from multiple columns
    locations = set()
    location_cols = ['City', 'Location', 'Headquarters', 'Australian_Office_Location', 'Office_Location']
    for col in location_cols:
        if col in df.columns:
            locations.update([x for x in df[col].dropna().unique() if str(x).strip() and str(x) != 'N/A'])
    locations = sorted(list(locations))
    
    # Generate market cap ranges based on actual data
    market_cap_ranges = [
        ('under_100m', 'Under $100M'),
        ('100m_500m', '$100M - $500M'), 
        ('500m_1b', '$500M - $1B'),
        ('1b_5b', '$1B - $5B'),
        ('over_5b', 'Over $5B')
    ]

    if request.method == "POST":
        # B. Get list of selected company indices
        selected = request.form.getlist("company")  # e.g. ["0", "3", "5"]
        if not selected:
            return "No companies selected", 400

        # Save the indices into a temp file for step 3
        with open("uploads/selected_companies.txt", "w") as f:
            f.write(",".join(selected))
        return redirect(url_for("generate_emails"))

    # If GET, render a page with checkboxes and filters
    return render_template("select.html", 
                         companies=companies_with_details,
                         exchanges=exchanges,
                         countries=countries,
                         sectors=sectors,
                         industries=industries,
                         locations=locations,
                         market_cap_ranges=market_cap_ranges,
                         current_filters={
                             'exchange': filter_exchange,
                             'country': filter_country,
                             'australian_sub': filter_australian_sub,
                             'sector': filter_sector,
                             'industry': filter_industry,
                             'market_cap_range': filter_market_cap,
                             'location': filter_location
                         })


# ----------------------------
# ROUTE #3: GENERATE EMAILS
# ----------------------------
@app.route("/generate", methods=["GET"])
def generate_emails():
    # A. Load DataFrame and selected indices
    df = pd.read_csv("uploads/last_upload.csv")
    try:
        with open("uploads/selected_companies.txt", "r") as f:
            selected_indices = [int(x) for x in f.read().split(",") if x.strip().isdigit()]
    except FileNotFoundError:
        return redirect(url_for("select_companies"))

    results = []  # list of dicts: { company, market_cap, headlines, email_draft }

    for idx in selected_indices:
        if idx < 0 or idx >= len(df):
            continue
        row = df.iloc[idx]
        company_name = str(row["Company"])
        
        # Extract comprehensive company information - handle missing columns gracefully
        market_cap = row.get("Market_Cap_USD", row.get("Market_Cap", row.get("MarketCap", "Unknown")))
        exchange = row.get("Exchange", row.get("Stock_Exchange", "Unknown"))
        sector = row.get("Sector", row.get("Industry", row.get("Business_Sector", "Unknown")))
        country = row.get("Country", row.get("Headquarters", row.get("Location", "Unknown")))
        has_aus_sub = row.get("Has_Australian_Subsidiary", row.get("Australian_Subsidiary", row.get("AUS_Subsidiary", "Unknown")))
        ceo_name = row.get("CEO_Name", row.get("CEO", row.get("Chief_Executive", "Unknown")))
        ceo_email = row.get("CEO_Email", row.get("CEO_Contact", "Unknown"))
        ir_contact = row.get("IR_Contact_Email", row.get("IR_Email", row.get("Investor_Relations", "Unknown")))

        # 1. Fetch recent news from NewsAPI
        news_url = (
            f"https://newsapi.org/v2/everything?"
            f"q={requests.utils.quote(company_name)}&"
            f"sortBy=publishedAt&"
            f"pageSize=5&"              # grab top 5 articles for more content
            f"apiKey={NEWSAPI_KEY}"
        )
        news_resp = requests.get(news_url)
        news_data = news_resp.json()
        headlines = []
        if news_data.get("status") == "ok":
            # Extract up to 5 headlines with descriptions
            for article in news_data.get("articles", [])[:5]:
                title = article.get("title", "")
                description = article.get("description", "")
                url = article.get("url", "")
                published_at = article.get("publishedAt", "")
                
                if title:
                    # Create rich content for better email personalization
                    if description and len(description) > 20:
                        headlines.append(f"• {title}\n  Summary: {description[:150]}{'...' if len(description) > 150 else ''}")
                    else:
                        headlines.append(f"• {title}")
        
        # Add fallback content if no news found
        if not headlines:
            headlines.append(f"Recent market activities and business developments for {company_name} indicate continued growth potential in their {sector.lower()} sector.")
        
        # Limit to top 3 most relevant headlines for the prompt
        headlines = headlines[:3]

        # 2. Call OpenAI to generate a personalized email draft
        # Build a more comprehensive prompt
        company_data = analyze_company_data(row, df.columns)
        prompt = create_intelligent_prompt(company_name, company_data, headlines)
        try:
            completion = client.chat.completions.create(
                model="gpt-3.5-turbo",  # free-trial tier model
                messages=[
                    {"role": "system", "content": "You are an email generator specializing in ASX listing outreach."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=250,
                temperature=0.7,
            )
            email_text = completion.choices[0].message.content.strip()
        except Exception as e:
            email_text = f"Error generating email: {e}"

        results.append({
            "company": company_name,
            "exchange": exchange,
            "sector": sector,
            "country": country,
            "market_cap": market_cap,
            "has_aus_sub": has_aus_sub,
            "ceo_name": ceo_name,
            "ceo_email": ceo_email,
            "ir_contact": ir_contact,
            "headlines": headlines,
            "email_draft": email_text,
            "analyzed_data": company_data  # Include the analyzed data for transparency
        })

    # 3. Render results in an HTML page
    return render_template("results.html", results=results)

def analyze_company_data(row, df_columns):
    """
    Intelligently analyze all available company data and categorize by relevance to ASX listing
    """
    company_data = {
        'financial_metrics': {},
        'business_info': {},
        'geographic_info': {},
        'leadership_info': {},
        'recent_activities': {},
        'other_relevant': {}
    }
    
    # Define column categories for ASX listing relevance
    financial_keywords = ['market_cap', 'revenue', 'profit', 'ebitda', 'assets', 'debt', 'cash', 'capital', 'valuation', 'shares', 'price', 'dividend']
    business_keywords = ['sector', 'industry', 'business', 'description', 'products', 'services', 'operations', 'subsidiaries']
    geographic_keywords = ['country', 'headquarters', 'location', 'region', 'australian', 'asia', 'pacific', 'office']
    leadership_keywords = ['ceo', 'cfo', 'chairman', 'director', 'executive', 'management', 'contact', 'email', 'ir', 'investor']
    activity_keywords = ['ipo', 'listing', 'raising', 'round', 'investment', 'acquisition', 'merger', 'expansion', 'growth']
    
    for column in df_columns:
        if column == 'Company':
            continue
            
        value = row.get(column, '')
        if pd.isna(value) or str(value).strip() == '' or str(value).lower() in ['unknown', 'n/a', 'na', 'none']:
            continue
            
        column_lower = column.lower()
        value_str = str(value)
        
        # Categorize by column name and content
        if any(keyword in column_lower for keyword in financial_keywords):
            company_data['financial_metrics'][column] = value_str
        elif any(keyword in column_lower for keyword in business_keywords):
            company_data['business_info'][column] = value_str
        elif any(keyword in column_lower for keyword in geographic_keywords):
            company_data['geographic_info'][column] = value_str
        elif any(keyword in column_lower for keyword in leadership_keywords):
            company_data['leadership_info'][column] = value_str
        elif any(keyword in column_lower for keyword in activity_keywords):
            company_data['recent_activities'][column] = value_str
        else:
            # Check if the content might be relevant
            if len(value_str) > 10 and any(word in value_str.lower() for word in ['million', 'billion', 'usd', 'aud', 'growth', 'market', 'international']):
                company_data['other_relevant'][column] = value_str
    
    return company_data

def create_intelligent_prompt(company_name, company_data, headlines):
    """
    Create a highly personalized prompt based on all available company data
    """
    
    # Extract key information
    financial_info = company_data['financial_metrics']
    business_info = company_data['business_info']
    geographic_info = company_data['geographic_info']
    leadership_info = company_data['leadership_info']
    activities_info = company_data['recent_activities']
    other_info = company_data['other_relevant']
    
    # Build dynamic sections based on available data
    prompt_sections = []
    
    prompt_sections.append(f"You are writing a highly personalized ASX listing proposal email to {company_name}.")
    
    # Company Profile Section
    profile_details = []
    if business_info:
        for key, value in business_info.items():
            if len(value) < 100:  # Avoid overly long descriptions
                profile_details.append(f"{key}: {value}")
    
    if financial_info:
        for key, value in financial_info.items():
            profile_details.append(f"{key}: {value}")
    
    if geographic_info:
        for key, value in geographic_info.items():
            profile_details.append(f"{key}: {value}")
    
    if profile_details:
        prompt_sections.append(f"COMPANY PROFILE:\n" + "\n".join(f"- {detail}" for detail in profile_details))
    
    # Leadership Information
    if leadership_info:
        leadership_details = []
        for key, value in leadership_info.items():
            leadership_details.append(f"{key}: {value}")
        prompt_sections.append(f"LEADERSHIP:\n" + "\n".join(f"- {detail}" for detail in leadership_details))
    
    # Recent Activities
    if activities_info:
        activity_details = []
        for key, value in activities_info.items():
            activity_details.append(f"{key}: {value}")
        prompt_sections.append(f"RECENT CORPORATE ACTIVITIES:\n" + "\n".join(f"- {detail}" for detail in activity_details))
    
    # News Analysis
    if headlines:
        prompt_sections.append(f"RECENT NEWS ANALYSIS:\n" + "\n".join(headlines))
    
    # Strategic recommendations based on available data
    strategic_points = []
    
    # Australian connection
    aus_connection = any('australian' in str(v).lower() for v in geographic_info.values())
    if aus_connection:
        strategic_points.append("- Leverage existing Australian presence for natural ASX progression")
    
    # Financial scale
    has_large_market_cap = any('billion' in str(v).lower() or 'million' in str(v).lower() for v in financial_info.values())
    if has_large_market_cap:
        strategic_points.append("- Scale and financial profile suitable for ASX institutional investors")
    
    # Geographic expansion
    asia_pacific_presence = any(region in str(geographic_info).lower() for region in ['asia', 'pacific', 'singapore', 'hong kong', 'japan'])
    if asia_pacific_presence:
        strategic_points.append("- Asia-Pacific presence aligns with ASX's regional investor base")
    
    # Add other relevant information
    if other_info:
        other_details = []
        for key, value in other_info.items():
            if len(value) < 150:  # Keep it concise
                other_details.append(f"{key}: {value}")
        if other_details:
            prompt_sections.append(f"ADDITIONAL RELEVANT INFORMATION:\n" + "\n".join(f"- {detail}" for detail in other_details))
    
    # Final instructions
    prompt_sections.append("""
TASK: Write a compelling, research-driven email that:
1. Opens with specific reference to their recent developments or business profile
2. Demonstrates deep understanding of their business and current situation
3. Connects their specific circumstances to ASX listing advantages
4. Creates urgency based on their growth trajectory and market position
5. Includes a clear call-to-action

STRATEGIC FOCUS AREAS:""")
    
    if strategic_points:
        prompt_sections.append("\n".join(strategic_points))
    else:
        prompt_sections.append("- Access to Australian and Asia-Pacific capital markets\n- Regulatory advantages and investor familiarity\n- Currency diversification benefits")
    
    prompt_sections.append("""
TONE: Professional, well-researched, and consultative (demonstrate you've done homework)
LENGTH: 175-225 words
FORMAT: Professional business email with clear next steps
REQUIREMENT: Make it clear this email is based on thorough research of their specific situation.""")
    
    return "\n\n".join(prompt_sections)

# For Vercel deployment
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8000) 