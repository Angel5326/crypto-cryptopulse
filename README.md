# 💰 CryptoPulse - Cryptocurrency Intelligence Platform

**CryptoPulse** is a real-time cryptocurrency intelligence platform that scrapes CoinMarketCap, analyzes market sentiment, tracks personal portfolio P&L, and exports professional Excel reports. It transforms raw market data into actionable insights for traders and investors.

---

## 📌 Features

| Feature | Description |
| :--- | :--- |
| **Live Market Data** | Scrapes Top 10 cryptocurrencies (Price, Market Cap, 24h Change, Volume). |
| **Sentiment Engine** | Calculates a dynamic **Fear/Greed Index** from 0–100 based on gainer ratios and average price movement. |
| **Portfolio Tracker** | Enter your holdings (e.g., `BTC: 0.01`) to see real-time USD valuation and allocation %. |
| **Smart Alerts** | Triggers **🚀 SURGE** (≥+5%) and **🔻 CRASH** (≤-5%) notifications. |
| **Professional Excel Export** | Generates a **4-sheet color-coded workbook** with charts, conditional formatting, and dominance maps. |
| **5 Output Files** | CSV history, session snapshot, portfolio history, Excel report, and text report. |

---

## 🛠️ Tech Stack

| Technology | Role |
| :--- | :--- |
| **Python 3.8+** | Core language for scraping, analysis, and orchestration |
| **Selenium & WebDriver Manager** | Scrapes dynamic CoinMarketCap tables with JavaScript rendering |
| **Pandas** | Data processing, CSV history management, and trend analysis |
| **OpenPyXL** | Creates beautifully formatted Excel reports with color-coding and charts |
| **Regex** | Parses abbreviated numbers (`$1.22T`, `$197B`, `2.4M`) into clean integers |

---

## 🧩 Architecture

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           CryptoPulse Architecture                              │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌─────────────────┐    ┌─────────────────┐    ┌─────────────────────────────┐ │
│  │   Selenium      │    │                 │    │                             │ │
│  │   WebDriver     │───▶│  CoinMarketCap  │───▶│  Raw Market Data            │ │
│  │   (Headless)    │    │  Homepage       │    │  Extraction                 │ │
│  └─────────────────┘    └─────────────────┘    └─────────────────────────────┘ │
│          │                                               │                     │
│          ▼                                               ▼                     │
│  ┌─────────────────┐    ┌─────────────────┐    ┌─────────────────────────────┐ │
│  │   Sentiment     │    │   Portfolio     │    │   5 Output Files            │ │
│  │   Engine        │───▶│   Tracker       │───▶│  (CSV + Excel + TXT)       │ │
│  │   (Fear/Greed)  │    │   (P&L)         │    │                             │ │
│  └─────────────────┘    └─────────────────┘    └─────────────────────────────┘ │
│                                                               │                 │
│                                                               ▼                 │
│  ┌─────────────────────────────────────────────────────────────────────────────┐│
│  │                         Terminal Dashboard                                 ││
│  │  • Market Table  • Spark Bars  • Movers  • Sentiment Gauge                ││
│  │  • Dominance Map  • Portfolio Snapshot  • Historical Trends               ││
│  └─────────────────────────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────────────────────────┘
```

---

## 📁 Project Structure

```
crypto-cryptopulse/
├── Crypto.py                 # Main intelligence platform
├── requirements.txt          # Dependencies (includes openpyxl)
├── .gitignore               # Excludes output files
└── output/                  # All generated reports
    ├── CryptoPulse_Report.xlsx   # 4-sheet Excel workbook
    ├── crypto_history.csv        # Appended log of every run
    ├── crypto_session_latest.csv # Current session snapshot
    ├── portfolio_history.csv     # Portfolio P&L log
    └── session_report.txt        # Plain-text summary log
```

---

## ⚙️ Installation & Setup

### Prerequisites
- Python 3.8 or higher
- Google Chrome browser (latest version)
- pip (Python package manager)

### Steps

**1. Clone the repository:**
```bash
git clone https://github.com/Angel5326/crypto-cryptopulse.git
cd crypto-cryptopulse
```

**2. Install dependencies:**
```bash
pip install -r requirements.txt
```

**`requirements.txt` contents:**
```
pandas>=2.0.0
selenium>=4.0.0
webdriver-manager>=4.0.0
openpyxl>=3.0.0
```

**3. (Optional) Edit Portfolio Holdings:**
Open `Crypto.py` and modify the `CONFIG['portfolio']` dictionary:
```python
"portfolio" : {
    "BTC"  : 0.01,    # 0.01 Bitcoin
    "ETH"  : 0.50,    # 0.50 Ethereum
    "BNB"  : 2.00,    # 2 BNB
    "SOL"  : 5.00,    # 5 Solana
}
```

**4. Run the tracker:**
```bash
python Crypto.py
```

---

## 🖥️ Usage

Once running, CryptoPulse will:

1. **Launch** Chrome in headless mode.
2. **Navigate** to CoinMarketCap homepage.
3. **Scrape** Top 10 coins (rank, name, symbol, price, change, market cap, volume).
4. **Enrich** data with Tier classification (MEGA, LARGE, MID, SMALL).
5. **Calculate** sentiment score (0–100) and label (EXTREME GREED → FEAR).
6. **Compute** your portfolio P&L in real-time.
7. **Display** a rich terminal dashboard with spark bars and movers.
8. **Export** 5 output files (CSV history, session snapshot, portfolio history, Excel report, text report).

### Sample Terminal Output:

```
  ╔══════════════════════════════════════════════════════════════╗
  ║           💼  PORTFOLIO  P&L  SNAPSHOT                     ║
  ╠══════════════════════════════════════════════════════════════╣
  ║  COIN         QTY       PRICE       VALUE USD    SHARE     ║
  ╠══════════════════════════════════════════════════════════════╣
  ║  BTC        0.0100    $62,124.02     $621.24      19.5%    ║
  ║  ETH        0.5000    $1,912.75      $956.38      30.0%    ║
  ║  BNB        2.0000    $600.72      $1,201.44      37.7%    ║
  ║  SOL        5.0000    $75.95         $379.75      11.9%    ║
  ╠══════════════════════════════════════════════════════════════╣
  ║  TOTAL PORTFOLIO VALUE                   $3,185.08          ║
  ╚══════════════════════════════════════════════════════════════╝
```

---

## 📸 Screenshots / Demo

> *(Add a screenshot of the terminal market table with prices and changes here)*

> *(Add a screenshot showing the 4 Excel sheets: Live Market, Portfolio, Dominance Map, History here)*

---

## 🧠 Engineering Decisions

### Why Selenium over Requests?
CoinMarketCap uses heavy JavaScript rendering for its tables. Selenium allows us to wait for dynamic content, handle lazy-loading, and execute JavaScript when needed.

### Why OpenPyXL for Excel Export?
OpenPyXL provides fine-grained control over cell formatting, conditional colors, and charts. It allows me to create professional, stakeholder-ready reports with minimal overhead.

### Why Classify Tiers (MEGA, LARGE, MID, SMALL)?
Tier classification provides quick visual context. MEGA coins (≥$200B) are blue-chip assets; SMALL coins (<$1B) are higher-risk. This mirrors how professional analysts categorize assets.

### Why Fear/Greed Index?
Market sentiment often drives price movements. The Fear/Greed index (0–100) provides a quantitative measure of market psychology, helping traders make contrarian decisions (e.g., buying when fear is high).

---

## 🧪 Testing

The script includes comprehensive error handling for:

- Network timeouts (retry logic with increasing delays)
- Missing or malformed HTML elements
- Parsing errors (handles T, B, M, K suffixes)
- UTF-8 encoding issues on Windows terminals

To run a test scrape:
```bash
python Crypto.py
```

Monitor the `output/` folder for generated files.

---

## 🚧 Limitations & Future Improvements

### Current Limitations
- **Top 10 only** – Currently scrapes only Top 10 coins. Could be extended to Top 100 or custom lists.
- **No scheduling** – Runs manually. Could be automated with cron jobs or GitHub Actions.
- **No cloud deployment** – Runs locally. Could be deployed as a serverless function.

### Future Improvements
- **Extended coin list** – Scrape Top 100 or allow custom coin selection.
- **Historical database** – Store results in a database (SQLite/PostgreSQL) for long-term trend analysis.
- **Web dashboard** – Build a web interface (Angular/React) to visualize trends interactively.
- **Alert system** – Send email/SMS alerts for significant price movements.
- **Technical indicators** – Add RSI, MACD, and moving averages for deeper analysis.

---

## 📬 Contact

**Abi Angelin** – [GitHub](https://github.com/Angel5326) [LinkedIn](https://www.linkedin.com/in/abiangelin)

---

*Built as part of a Python Internship at Cybernaut.*
