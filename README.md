# 🚀 Stock Exchange Email Generator

A powerful AI-driven web application that transforms company databases into personalized stock exchange listing outreach emails. Built with Flask, powered by OpenAI GPT, and enhanced with real-time news integration.

![Stock Exchange Email Generator](https://img.shields.io/badge/Status-Live-brightgreen)
![Python](https://img.shields.io/badge/Python-3.8+-blue)
![Flask](https://img.shields.io/badge/Flask-3.1.1-green)
![OpenAI](https://img.shields.io/badge/OpenAI-GPT--3.5-orange)

## ✨ Features

### 🧠 **AI-Powered Email Generation**
- **GPT-3.5 Integration**: Creates highly personalized, professional emails
- **Context-Aware**: Analyzes company data to craft relevant messaging
- **Dynamic Content**: Incorporates company-specific details and market context

### 📊 **Intelligent Data Processing**
- **Smart Column Detection**: Automatically maps various CSV column formats
- **Robust CSV Parsing**: Handles malformed files with multiple fallback strategies
- **Data Categorization**: Intelligently sorts company information by relevance

### 📰 **Real-Time News Integration**
- **NewsAPI Integration**: Fetches latest company-specific headlines
- **Contextual Analysis**: Connects news to stock exchange listing opportunities
- **Rich Content**: Includes article summaries for deeper insights

### 🔍 **Advanced Filtering System**
- **Dynamic Filters**: Market cap ranges, geographic locations, industry sectors
- **Multi-Column Search**: Searches across various location and industry fields
- **Real-Time Updates**: Instant filtering without page reloads

### 🎨 **Modern UI/UX**
- **Responsive Design**: Works seamlessly on desktop and mobile
- **Glass Morphism**: Modern backdrop-blur effects and gradients
- **Interactive Elements**: Smooth animations and micro-interactions
- **Accessibility**: Keyboard navigation and screen reader support

## 🛠 Technology Stack

- **Backend**: Flask 3.1.1, Python 3.8+
- **AI/ML**: OpenAI GPT-3.5-turbo
- **Data Processing**: Pandas, NumPy
- **News API**: NewsAPI for real-time headlines
- **Frontend**: Modern CSS3, Vanilla JavaScript
- **Deployment**: Vercel (serverless)

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- OpenAI API key
- NewsAPI key

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/nicholasconoplia/miningmetals.git
   cd miningmetals
   ```

2. **Create virtual environment**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Set up environment variables**
   ```bash
   # Create .env file
   echo "OPENAI_API_KEY=your_openai_api_key_here" > .env
   echo "NEWSAPI_KEY=your_newsapi_key_here" >> .env
   ```

5. **Run the application**
   ```bash
   python main.py
   ```

6. **Open your browser**
   Navigate to `http://localhost:8000`

## 📋 Usage Guide

### Step 1: Upload Company Data
- **Supported Formats**: CSV, Excel (.xlsx, .xls)
- **Required Column**: Company name (various formats supported)
- **Optional Columns**: Exchange, Sector, Country, Market Cap, etc.

### Step 2: Filter & Select Companies
- **Apply Filters**: Use dynamic filters to narrow down companies
- **Select Companies**: Choose specific companies or select all
- **Preview Data**: Review company details before processing

### Step 3: Generate Emails
- **AI Processing**: GPT analyzes each company's profile
- **News Integration**: Real-time headlines incorporated
- **Email Generation**: Personalized emails created instantly

### Step 4: Review & Export
- **Copy Individual Emails**: One-click copy to clipboard
- **Download All**: Export all emails as text file
- **Review Analysis**: See what data influenced each email

## 🔧 Configuration

### API Keys Setup

1. **OpenAI API Key**
   - Visit [OpenAI Platform](https://platform.openai.com)
   - Create account and add billing ($5 minimum)
   - Generate API key in API section

2. **NewsAPI Key**
   - Visit [NewsAPI](https://newsapi.org)
   - Sign up for free account (100 requests/day)
   - Verify email and get API key

### Environment Variables
```bash
OPENAI_API_KEY=sk-your-openai-key-here
NEWSAPI_KEY=your-newsapi-key-here
```

## 📁 Project Structure

```
miningmetals/
├── templates/
│   ├── upload.html      # File upload interface
│   ├── select.html      # Company selection & filtering
│   └── results.html     # Generated emails display
├── uploads/             # Temporary file storage
├── main.py             # Main Flask application
├── app.py              # Vercel deployment entry
├── requirements.txt    # Python dependencies
├── vercel.json         # Vercel configuration
├── .env               # Environment variables
├── .gitignore         # Git ignore rules
└── README.md          # Project documentation
```

## 🌐 Deployment

### Deploy to Vercel

1. **Connect to GitHub**
   ```bash
   git remote add origin https://github.com/nicholasconoplia/miningmetals.git
   git branch -M main
   git push -u origin main
   ```

2. **Set up Vercel**
   - Import project from GitHub
   - Add environment variables in Vercel dashboard
   - Deploy automatically

3. **Environment Variables in Vercel**
   ```
   NEWSAPI_KEY: your_newsapi_key
   OPENAI_API_KEY: your_openai_key
   ```

### Local Development

```bash
# Run with debug mode
python main.py

# Access at http://localhost:8000
```

## 📊 Supported Data Formats

### Required Columns (any of these names):
- Company, Company Name, Name, Symbol, Ticker

### Optional Columns:
- **Exchange**: NYSE, NASDAQ, TSX, LSE, etc.
- **Geographic**: Country, City, Location, Headquarters
- **Financial**: Market Cap, Revenue, EBITDA
- **Industry**: Sector, Industry, Business Type
- **Contacts**: CEO, Email, Investor Relations

## 🔒 Security & Privacy

- **No Data Storage**: Uploaded files are temporary and auto-deleted
- **API Security**: Environment variables protect sensitive keys
- **Privacy First**: No personal data retained after session

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support & Issues

- **Issues**: [GitHub Issues](https://github.com/nicholasconoplia/miningmetals/issues)
- **Documentation**: This README
- **Contact**: Open an issue for support

## 🙏 Acknowledgments

- **OpenAI**: For GPT-3.5 API
- **NewsAPI**: For real-time news data
- **Flask Community**: For the excellent framework
- **Contributors**: All who helped improve this project

---

Built with ❤️ for the Australian Securities Exchange ecosystem 