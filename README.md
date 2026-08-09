# 💰 CryptoPulse - Cryptocurrency Intelligence Platform

**A real-time cryptocurrency tracker that scrapes CoinMarketCap, analyzes market sentiment, and tracks your personal portfolio value.**

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Selenium](https://img.shields.io/badge/Selenium-WebDriver-green.svg)
![Excel](https://img.shields.io/badge/Excel-OpenPyXL-brightgreen.svg)

---

## 📌 Features

- **Live Market Data** – Scrapes Top 10 Cryptocurrencies (Price, Market Cap, 24h Change, Volume).
- **Sentiment Engine** – Calculates a dynamic **Fear/Greed Index** from 0–100 based on gainer ratios and average price movement.
- **Portfolio Tracker** – Enter your holdings in `CONFIG['portfolio']` (e.g., `BTC: 0.01`) to see your real-time P&L in USD.
- **Smart Alerts** – Triggers **🚀 SURGE** and **🔻 CRASH** notifications for movements exceeding ±5%.
- **Professional Excel Export** – Generates a 4-sheet workbook with color-coded cells, charts, and market dominance maps.

---

## 🛠️ Tech Stack

- **Python 3.8+**
- **Selenium & WebDriver Manager** – Scrapes dynamic CoinMarketCap tables.
- **Pandas** – Data processing and CSV history management.
- **OpenPyXL** – Creates beautifully formatted Excel reports.
- **Terminal Colors** – ANSI-coded rich terminal dashboard.

---

## 🚀 Getting Started

### 1. Clone the Repository
```bash
git clone https://github.com/Angel5326/crypto-cryptopulse.git
cd crypto-cryptopulse