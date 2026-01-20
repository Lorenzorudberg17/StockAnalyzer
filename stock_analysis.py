"""
Stock Analysis Tool for Excel 
Requirements: pip install xlwings yfinance pandas matplotlib requests beautifulsoup4
"""

import xlwings as xw
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup

def get_stock_data(ticker, period):
    """Fetch stock data from Yahoo Finance"""
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period=period)
        info = stock.info
        return stock, hist, info
    except Exception as e:
        return None, None, None
    

    
def get_statement_metrics(stock):
    """
    Pull Income Statement + Cash Flow items (TTM-ish using latest annual column).
    Returns raw numbers.
    """
    out = {
        "Revenue (TTM)": None,
        "Net Income (TTM)": None,
        "Operating Income (TTM)": None,
        "EBITDA (TTM)": None,
        "Cash From Ops (TTM)": None,
        "CapEx (TTM)": None,
        "Free Cash Flow (TTM)": None,
    }

    # Income Statement
    try:
        fin = stock.financials
        if fin is not None and not fin.empty:
            col = fin.columns[0]
            if "Total Revenue" in fin.index:
                out["Revenue (TTM)"] = fin.loc["Total Revenue", col]
            if "Net Income" in fin.index:
                out["Net Income (TTM)"] = fin.loc["Net Income", col]
            if "Operating Income" in fin.index:
                out["Operating Income (TTM)"] = fin.loc["Operating Income", col]
            if "EBITDA" in fin.index:
                out["EBITDA (TTM)"] = fin.loc["EBITDA", col]
    except:
        pass

    # Cash Flow
    try:
        cf = stock.cashflow
        if cf is not None and not cf.empty:
            col = cf.columns[0]
            if "Total Cash From Operating Activities" in cf.index:
                out["Cash From Ops (TTM)"] = cf.loc["Total Cash From Operating Activities", col]
            if "Capital Expenditures" in cf.index:
                out["CapEx (TTM)"] = cf.loc["Capital Expenditures", col]

            cfo = out["Cash From Ops (TTM)"]
            capex = out["CapEx (TTM)"]
            if cfo is not None and capex is not None:
                out["Free Cash Flow (TTM)"] = cfo + capex  # capex normally negative
    except:
        pass

    return out


def fmt_money(x):
    if x is None:
        return "N/A"
    try:
        x = float(x)
        sign = "-" if x < 0 else ""
        x = abs(x)

        if x >= 1e12:
            return f"{sign}${x/1e12:.2f}T"
        if x >= 1e9:
            return f"{sign}${x/1e9:.2f}B"
        if x >= 1e6:
            return f"{sign}${x/1e6:.2f}M"
        return f"{sign}${x:,.0f}"
    except:
        return "N/A"
    


def format_metrics_table(sheet, start_row, end_row):
    """Apply professional formatting to metrics table"""
    # Header styling - dark blue background, white text
    header_cells = [start_row + i for i in [0, 5, 9, 11, 16, 18, 19]]  # Section headers
    for row in header_cells:
        cell_range = sheet.range(f'A{row}:B{row}')
        cell_range.color = (31, 78, 120)  # Dark blue
        cell_range.api.Font.Color = 0xFFFFFF  # White text
        cell_range.api.Font.Bold = True
        cell_range.api.Font.Size = 12
    
    # Metric rows - alternating colors
    metric_row = start_row + 1
    while metric_row <= end_row:
        if sheet.range(f'A{metric_row}').value and sheet.range(f'A{metric_row}').api.Font.Bold == False:
            if (metric_row - start_row) % 2 == 0:
                sheet.range(f'A{metric_row}:B{metric_row}').color = (242, 242, 242)  # Light gray
            else:
                sheet.range(f'A{metric_row}:B{metric_row}').color = (255, 255, 255)  # White
        metric_row += 1
    
    # Add borders
    for row in range(start_row, end_row + 1):
        cell_range = sheet.range(f'A{row}:B{row}')
        # Add thin borders
        for border_id in [7, 8, 9, 10]:  # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
            cell_range.api.Borders(border_id).LineStyle = 1
            cell_range.api.Borders(border_id).Weight = 2

def calculate_key_metrics(info, hist, statement_metrics):
    """
    Return metrics in clean grouped order:
    Income Statement -> Profitability -> Growth -> Valuation -> Dividends -> Risk
    """
    metrics = {}

    try:
        # =========================
        # INCOME STATEMENT
        # =========================
        metrics["Revenue (TTM)"] = fmt_money(statement_metrics.get("Revenue (TTM)"))
        metrics["Net Income (TTM)"] = fmt_money(statement_metrics.get("Net Income (TTM)"))
        metrics["Operating Income (TTM)"] = fmt_money(statement_metrics.get("Operating Income (TTM)"))
        metrics["EBITDA (TTM)"] = fmt_money(statement_metrics.get("EBITDA (TTM)"))

        # =========================
        # PROFITABILITY & MARGINS
        # =========================
        revenue = statement_metrics.get("Revenue (TTM)")
        net_income = statement_metrics.get("Net Income (TTM)")
        op_inc = statement_metrics.get("Operating Income (TTM)")
        ebitda = statement_metrics.get("EBITDA (TTM)")

        # Profit Margin
        if revenue and revenue != 0 and net_income is not None:
            metrics["Profit Margin"] = f"{(net_income / revenue) * 100:.2f}%"
        else:
            metrics["Profit Margin"] = "N/A"

        # Operating Margin
        if revenue and revenue != 0 and op_inc is not None:
            metrics["Operating Margin"] = f"{(op_inc / revenue) * 100:.2f}%"
        else:
            metrics["Operating Margin"] = "N/A"

        # EBITDA Margin
        if revenue and revenue != 0 and ebitda is not None:
            metrics["EBITDA Margin"] = f"{(ebitda / revenue) * 100:.2f}%"
        else:
            metrics["EBITDA Margin"] = "N/A"

        # ROE (Return on Equity)
        roe = info.get("returnOnEquity", None)
        if roe is not None:
            metrics["Return on Equity (ROE)"] = f"{roe * 100:.2f}%"
        else:
            metrics["Return on Equity (ROE)"] = "N/A"

        # =========================
        # GROWTH
        # =========================
        # Revenue Growth
        rev_growth = info.get("revenueGrowth", None)
        if rev_growth is not None:
            metrics["Revenue Growth (YoY)"] = f"{rev_growth * 100:.2f}%"
        else:
            metrics["Revenue Growth (YoY)"] = "N/A"

        # Earnings Growth
        earnings_growth = info.get("earningsGrowth", None)
        if earnings_growth is not None:
            metrics["Earnings Growth (YoY)"] = f"{earnings_growth * 100:.2f}%"
        else:
            metrics["Earnings Growth (YoY)"] = "N/A"

        # =========================
        # VALUATION
        # =========================
        metrics["Current Price"] = info.get("currentPrice", "N/A")
        metrics["Market Cap"] = fmt_money(info.get("marketCap", None))
        # 52-Week Range
        fifty_two_week_high = info.get("fiftyTwoWeekHigh", None)
        fifty_two_week_low = info.get("fiftyTwoWeekLow", None)
        
        metrics["52-Week High"] = f"${fifty_two_week_high:.2f}" if fifty_two_week_high else "N/A"
        metrics["52-Week Low"] = f"${fifty_two_week_low:.2f}" if fifty_two_week_low else "N/A"
        
        # Calculate distance from 52-week high
        current_price = info.get("currentPrice", None)
        if current_price and fifty_two_week_high:
            distance_from_high = ((current_price - fifty_two_week_high) / fifty_two_week_high) * 100
            metrics["Distance from 52W High"] = f"{distance_from_high:.2f}%"
        else:
            metrics["Distance from 52W High"] = "N/A"
        trailing_pe = info.get("trailingPE", None)
        forward_pe = info.get("forwardPE", None)
        ps_ratio = info.get("priceToSalesTrailing12Months", None)
        
        metrics["P/E Ratio"] = round(trailing_pe, 2) if trailing_pe else "N/A"
        metrics["Forward P/E"] = round(forward_pe, 2) if forward_pe else "N/A"
        metrics["Price to Sales (P/S)"] = round(ps_ratio, 2) if ps_ratio else "N/A"

        # =========================
        # DIVIDENDS
        # =========================
        div_yield = info.get("dividendYield", None)
        if div_yield is not None:
            metrics["Dividend Yield"] = f"{div_yield * 100:.2f}%"
        else:
            metrics["Dividend Yield"] = "N/A"

        payout_ratio = info.get("payoutRatio", None)
        if payout_ratio is not None:
            metrics["Payout Ratio"] = f"{payout_ratio * 100:.2f}%"
        else:
            metrics["Payout Ratio"] = "N/A"

        # =========================
        # RISK
        # =========================
        metrics["Beta"] = round(info.get("beta", 0), 2) if info.get("beta") else "N/A"

    except Exception as e:
        print(f"Error calculating metrics: {e}")

    return metrics

def get_stock_news(ticker):
    """Fetch recent news for the stock"""
    news_data = []
    
    # Method 1: Try yfinance news
    try:
        stock = yf.Ticker(ticker)
        news = stock.news
        
        if news and len(news) > 0:
            for item in news[:10]:
                title = item.get('title', 'N/A')
                publisher = item.get('publisher', 'N/A')
                link = item.get('link', 'N/A')
                
                pub_time = item.get('providerPublishTime', 0)
                if pub_time and pub_time > 0:
                    try:
                        published = datetime.fromtimestamp(pub_time).strftime('%Y-%m-%d %H:%M')
                    except:
                        published = 'Recent'
                else:
                    published = 'Recent'
                if title == 'N/A' and publisher == 'N/A' and link == 'N/A':
                    continue
                news_data.append({
                    'Title': title,
                    'Publisher': publisher,
                    'Link': link,
                    'Published': published
                })
            
            if len(news_data) > 0:
                return news_data
    except Exception as e:
        print(f"yfinance news failed: {e}")
    
    # Method 2: Google News RSS fallback
    try:
        rss_url = f"https://news.google.com/rss/search?q={ticker}+stock&hl=en-US&gl=US&ceid=US:en"
        response = requests.get(rss_url, timeout=10)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'xml')
            items = soup.find_all('item')[:10]
            
            for item in items:
                title = item.find('title').text if item.find('title') else 'N/A'
                link = item.find('link').text if item.find('link') else 'N/A'
                pub_date = item.find('pubDate').text if item.find('pubDate') else 'Recent'
                
                if ' - ' in title:
                    parts = title.rsplit(' - ', 1)
                    actual_title = parts[0]
                    publisher = parts[1] if len(parts) > 1 else 'Google News'
                else:
                    actual_title = title
                    publisher = 'Google News'
                
                news_data.append({
                    'Title': actual_title,
                    'Publisher': publisher,
                    'Link': link,
                    'Published': pub_date[:16] if len(pub_date) > 16 else pub_date
                })
            
            if len(news_data) > 0:
                return news_data
    except Exception as e:
        print(f"Google RSS failed: {e}")
    
    # Fallback
    if len(news_data) == 0:
        news_data.append({
            'Title': f'Click to view {ticker} news on Yahoo Finance',
            'Publisher': 'Yahoo Finance',
            'Link': f'https://finance.yahoo.com/quote/{ticker}/news',
            'Published': 'N/A'
        })
    
    return news_data

def update_news_sheet(ticker, wb):
    """Update news sheet with latest news"""
    news_sheet = wb.sheets['News']
    news_sheet.clear_contents()

    news_data = get_stock_news(ticker)

    news_sheet.range('A1').value = f"Recent News for {ticker}"
    news_sheet.range('A1').font.size = 16
    news_sheet.range('A1').font.bold = True

    news_sheet.range('A3').value = [['Title', 'Publisher', 'Published', 'Link']]
    news_sheet.range('A3:D3').font.bold = True

    row = 4
    for news in (news_data or []):
        news_sheet.range(f'A{row}').value = news.get('Title', '')
        news_sheet.range(f'B{row}').value = news.get('Publisher', '')
        news_sheet.range(f'C{row}').value = news.get('Published', '')

        link = news.get('Link', '')
        cell = news_sheet.range(f'D{row}')
        cell.value = ""

        if isinstance(link, str) and link.startswith("http"):
            try:
                # Reliable hyperlink add (avoids #VALUE! formula issues)
                cell.api.Hyperlinks.Add(Anchor=cell.api, Address=link, TextToDisplay="View Article")
            except:
                # fallback: show raw URL
                cell.value = link
        else:
            cell.value = ""

        row += 1

    # Format columns
    news_sheet.range('A:A').column_width = 60
    news_sheet.range('B:B').column_width = 20
    news_sheet.range('C:C').column_width = 18
    news_sheet.range('D:D').column_width = 14



def analyze_stock():
    wb = xw.Book.caller()
    single_sheet = wb.sheets['Single']

    # --- Inputs ---
    ticker = single_sheet.range('B2').value
    period_option = single_sheet.range('B3').value

    if not ticker:
        single_sheet.range('B5').value = "Please enter a ticker symbol"
        return

    ticker = str(ticker).upper().strip()

    # Map dropdown text -> yfinance period codes
    period_map = {
        '1 Week': '5d',
        '1 Month': '1mo',
        '3 Months': '3mo',
        '6 Months': '6mo',
        '1 Year': '1y',
        '2 Years': '2y',
        '5 Years': '5y',
        'Max': 'max'
    }
    period = period_map.get(str(period_option).strip(), '1y')

    # Status
    single_sheet.range('B5').value = f"Loading data for {ticker}..."

    # --- Fetch data ---
    stock, hist, info = get_stock_data(ticker, period)
    if stock is None or hist is None or hist.empty or info is None:
        single_sheet.range('B5').value = f"Error: Could not fetch data for {ticker}"
        return

    # THIS IS WHERE THE INDENTATION WAS WRONG - this try block needs to be at this level
    try:
        # Remove old chart if it exists
        try:
            single_sheet.pictures['StockChart'].delete()
        except Exception:
            pass

        # Clear only the table/text area (safe)
        single_sheet.range('A7:C200').clear_contents()

    except Exception as e:
        print(f"Warning: could not fully clear prior output: {e}")

    # --- Metrics ---
    statement_metrics = get_statement_metrics(stock)
    metrics = calculate_key_metrics(info, hist, statement_metrics)

    # === SECTION 1: Company Header ===
    row = 7
    # ... rest of your code


    single_sheet.range(f'A{row}').value = f"STOCK ANALYSIS: {ticker}"
    single_sheet.range(f'A{row}').font.size = 18
    single_sheet.range(f'A{row}').font.bold = True
    row += 1

    single_sheet.range(f'A{row}').value = f"Company: {info.get('longName', ticker)}"
    single_sheet.range(f'A{row}').font.size = 12
    row += 1

    single_sheet.range(f'A{row}').value = (
        f"Sector: {info.get('sector', 'N/A')} | Industry: {info.get('industry', 'N/A')}"
    )
    row += 1

    single_sheet.range(f'A{row}').value = f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    row += 2

    # === SECTION 2: Key Metrics Table ===
    single_sheet.range(f'A{row}').value = "KEY METRICS"
    single_sheet.range(f'A{row}').font.bold = True
    single_sheet.range(f'A{row}').font.size = 16
    single_sheet.range(f'A{row}:B{row}').color = (68, 114, 196)  # Blue header
    single_sheet.range(f'A{row}').api.Font.Color = 0xFFFFFF  # White text
    row += 2

    start_metrics_row = row

    start_metrics_row = row

    # Define sections (names must match keys in `metrics`)
    sections = {
        "INCOME STATEMENT": [
            "Revenue (TTM)",
            "Net Income (TTM)",
            "Operating Income (TTM)",
            "EBITDA (TTM)",
        ],
        "PROFITABILITY & MARGINS": [
            "Profit Margin",
            "Operating Margin",
            "EBITDA Margin",
            "Return on Equity (ROE)",
        ],
        "GROWTH": [
            "Revenue Growth (YoY)",
            "Earnings Growth (YoY)",
        ],
        "VALUATION": [
            "Current Price",
            "Market Cap",
            "52-Week High",
            "52-Week Low",
            "Distance from 52W High",
            "P/E Ratio",
            "Forward P/E",
            "Price to Sales (P/S)",
        ],
        "DIVIDENDS": [
            "Dividend Yield",
            "Payout Ratio",
        ],
        "RISK": [
            "Beta",
        ],
    }

    for section_name, keys in sections.items():
        # Section Header
        single_sheet.range(f"A{row}").value = section_name
        single_sheet.range(f"A{row}:B{row}").color = (31, 78, 120)  # Dark blue
        single_sheet.range(f"A{row}").api.Font.Color = 0xFFFFFF  # White
        single_sheet.range(f"A{row}").font.bold = True
        single_sheet.range(f"A{row}").font.size = 11
        row += 1

        # Rows under section with alternating colors
        for i, key in enumerate(keys):
            single_sheet.range(f"A{row}").value = key
            single_sheet.range(f"B{row}").value = metrics.get(key, "N/A")
            
            # Alternating row colors
            if i % 2 == 0:
                single_sheet.range(f'A{row}:B{row}').color = (242, 242, 242)  # Light gray
            else:
                single_sheet.range(f'A{row}:B{row}').color = (255, 255, 255)  # White
            
            # Right align numbers
            single_sheet.range(f"B{row}").api.HorizontalAlignment = -4152  # xlRight
            
            row += 1

        row += 1  # blank line after each section

    # Add borders to entire table
    table_range = single_sheet.range(f'A{start_metrics_row}:B{row-1}')
    for border_id in [7, 8, 9, 10, 11, 12]:  # All border types
        table_range.api.Borders(border_id).LineStyle = 1
        table_range.api.Borders(border_id).Weight = 2
        table_range.api.Borders(border_id).Color = 0x000000

    # Format columns
    single_sheet.range('A:A').column_width = 24
    single_sheet.range('B:B').column_width = 20
    # === SECTION 3: Enhanced Stock Price Chart with Volume ===
    try:
        # Create figure with 2 subplots (price and volume)
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8), 
                                        gridspec_kw={'height_ratios': [3, 1]}, 
                                        sharex=True)
        
        # ========== TOP CHART: PRICE + MOVING AVERAGES ==========
        # Plot closing price
        ax1.plot(hist.index, hist['Close'], linewidth=2.5, color='#2E86DE', label=f'{ticker}', zorder=3)
        
        # Calculate and plot moving averages
        if len(hist) >= 50:
            ma50 = hist['Close'].rolling(window=50).mean()
            ax1.plot(hist.index, ma50, linewidth=1.5, color='#F39C12', 
                    label='50-day MA', linestyle='--', alpha=0.8, zorder=2)
        
        if len(hist) >= 200:
            ma200 = hist['Close'].rolling(window=200).mean()
            ax1.plot(hist.index, ma200, linewidth=1.5, color='#E74C3C', 
                    label='200-day MA', linestyle='--', alpha=0.8, zorder=2)
        
        # Add 52-week high/low lines
        fifty_two_week_high = info.get("fiftyTwoWeekHigh", None)
        fifty_two_week_low = info.get("fiftyTwoWeekLow", None)
        
        if fifty_two_week_high:
            ax1.axhline(y=fifty_two_week_high, color='#27AE60', linestyle=':', 
                       linewidth=1.5, alpha=0.6, label='52W High', zorder=1)
        
        if fifty_two_week_low:
            ax1.axhline(y=fifty_two_week_low, color='#C0392B', linestyle=':', 
                       linewidth=1.5, alpha=0.6, label='52W Low', zorder=1)
        
        # Add current price annotation
        current_price = hist['Close'].iloc[-1]
        last_date = hist.index[-1]
        ax1.annotate(f'${current_price:.2f}', 
                    xy=(last_date, current_price),
                    xytext=(10, 0), 
                    textcoords='offset points',
                    fontsize=10,
                    fontweight='bold',
                    color='#2E86DE',
                    bbox=dict(boxstyle='round,pad=0.4', facecolor='white', 
                             edgecolor='#2E86DE', linewidth=2),
                    zorder=5)
        
        # Add dot at current price
        ax1.scatter([last_date], [current_price], color='#2E86DE', s=100, zorder=5)
        
        # Calculate price change
        start_price = hist['Close'].iloc[0]
        price_change = current_price - start_price
        price_change_pct = (price_change / start_price) * 100
        change_color = '#27AE60' if price_change >= 0 else '#E74C3C'
        change_sign = '+' if price_change >= 0 else ''
        
        # Title with price change
        title_text = f'{ticker} Stock Price - {period_option}\n{change_sign}${price_change:.2f} ({change_sign}{price_change_pct:.2f}%)'
        ax1.set_title(title_text, fontsize=15, fontweight='bold', pad=15, color=change_color)
        
        ax1.set_ylabel('Price ($)', fontsize=11, fontweight='bold')
        ax1.grid(True, alpha=0.2, linestyle='--')
        ax1.set_facecolor('#F8F9FA')
        ax1.legend(loc='upper left', fontsize=9, framealpha=0.9)
        
        # ========== BOTTOM CHART: VOLUME ==========
        colors = ['#27AE60' if hist['Close'].iloc[i] >= hist['Open'].iloc[i] 
                  else '#E74C3C' for i in range(len(hist))]
        
        ax2.bar(hist.index, hist['Volume'], color=colors, alpha=0.6, width=0.8)
        ax2.set_ylabel('Volume', fontsize=11, fontweight='bold')
        ax2.set_xlabel('Date', fontsize=11, fontweight='bold')
        ax2.grid(True, alpha=0.2, linestyle='--', axis='y')
        ax2.set_facecolor('#F8F9FA')
        
        # Format volume numbers (millions/billions)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{x/1e6:.0f}M' if x >= 1e6 else f'{x:.0f}'))
        
        # Format x-axis dates as MM/YY
        import matplotlib.dates as mdates
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%m/%y'))
        ax2.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        
        plt.xticks(rotation=45, ha='right')
        
        fig.patch.set_facecolor('white')
        plt.tight_layout()
        
        # Add chart to Excel
        single_sheet.pictures.add(fig, name='StockChart', update=True,
                                   left=single_sheet.range('E7').left,
                                   top=single_sheet.range('E7').top)
        plt.close()
    except Exception as e:
        print(f"Error creating chart: {e}")
        single_sheet.range('E7').value = f"Chart error: {e}"

    # === Update News Sheet ===
    try:
        update_news_sheet(ticker, wb)
    except Exception as e:
        print(f"Error updating news: {e}")

    # Success message
    try:
        single_sheet.range('B5').value = f"Analysis complete for {ticker}!"
        single_sheet.range('B5').font.color = (0, 128, 0)
    except Exception as e:
        single_sheet.range('B5').value = f"Error at end: {str(e)}"
        print(f"Final error: {e}")













def compare_stocks():
    """Compare multiple stocks side by side"""
    try:
        # Get user inputs
        compare_sheet = xw.sheets['Compare']
        
        # Read tickers from individual cells (B2, B3, B4)
        ticker1 = compare_sheet.range('B2').value
        ticker2 = compare_sheet.range('B3').value
        ticker3 = compare_sheet.range('B4').value
        
        period_map = {
            '1 Week': '5d',
            '1 Month': '1mo',
            '3 Months': '3mo',
            '6 Months': '6mo',
            '1 Year': '1y',
            '2 Years': '2y',
            '5 Years': '5y',
            'Max': 'max'
        }
        
        period_option = compare_sheet.range('B5').value or '1 Year'
        period = period_map.get(str(period_option).strip(), '1y')
        
        # Build tickers list from individual cells
        tickers = []
        for t in [ticker1, ticker2, ticker3]:
            if t and str(t).strip():
                tickers.append(str(t).strip().upper())
        
        if len(tickers) < 2:
            compare_sheet.range('B7').value = "Please enter at least 2 tickers"
            return
        
        # Clear previous
        compare_sheet.range('A10:Z200').clear_contents()

        # Fetch all data
        stock_data = {}
        for i, ticker in enumerate(tickers):
            compare_sheet.range('B7').value = f"Loading {ticker} ({i+1}/{len(tickers)})..."
            
            stock, hist, info = get_stock_data(ticker, period)
            if stock and hist is not None and not hist.empty:
                try:
                    statement_metrics = get_statement_metrics(stock)
                    metrics = calculate_key_metrics(info, hist, statement_metrics)
                    stock_data[ticker] = {'stock': stock, 'hist': hist, 'info': info, 'metrics': metrics}
                except Exception as e:
                    compare_sheet.range('B7').value = f"Error processing {ticker}: {str(e)}"
                    print(f"Error with {ticker}: {e}")
                    return
            else:
                compare_sheet.range('B7').value = f"Error loading {ticker} - no data returned"
                return
            
        # === Comparison Table ===
        compare_sheet.range('B7').value = f"Building comparison table..."
        
        row = 10
        compare_sheet.range(f'A{row}').value = "STOCK COMPARISON"
        compare_sheet.range(f'A{row}').font.bold = True
        compare_sheet.range(f'A{row}').font.size = 16
        
        # Color the main header
        header_range = compare_sheet.range(f'A{row}:{chr(65 + len(tickers))}{row}')
        header_range.color = (68, 114, 196)  # Blue
        compare_sheet.range(f'A{row}').api.Font.Color = 0xFFFFFF
        row += 2
        
        start_table_row = row
        
        # Build column headers
        compare_sheet.range(f'A{row}').value = "Metric"
        compare_sheet.range(f'A{row}').font.bold = True
        compare_sheet.range(f'A{row}').color = (31, 78, 120)
        compare_sheet.range(f'A{row}').api.Font.Color = 0xFFFFFF
        
        for i, ticker in enumerate(tickers):
            cell = compare_sheet.range(row, i + 2)
            cell.value = ticker
            cell.font.bold = True
            cell.color = (31, 78, 120)
            cell.api.Font.Color = 0xFFFFFF
            cell.api.HorizontalAlignment = -4108  # Center
        row += 1
        
        # Company names
        compare_sheet.range(f'A{row}').value = "Company Name"
        compare_sheet.range(f'A{row}').font.italic = True
        compare_sheet.range(f'A{row}').color = (217, 217, 217)
        
        for i, ticker in enumerate(tickers):
            cell = compare_sheet.range(row, i + 2)
            cell.value = stock_data[ticker]['info'].get('longName', ticker)
            cell.font.size = 9
            cell.color = (217, 217, 217)
            cell.api.HorizontalAlignment = -4108
        row += 1
        
        # Define sections
        sections = {
            "INCOME STATEMENT": [
                "Revenue (TTM)", "Net Income (TTM)", "Operating Income (TTM)", "EBITDA (TTM)",
            ],
            "PROFITABILITY & MARGINS": [
                "Profit Margin", "Operating Margin", "EBITDA Margin", "Return on Equity (ROE)",
            ],
            "GROWTH": [
                "Revenue Growth (YoY)", "Earnings Growth (YoY)",
            ],
            "VALUATION": [
                "Current Price", "Market Cap", "52-Week High", "52-Week Low", 
                "Distance from 52W High", "P/E Ratio", "Forward P/E", "Price to Sales (P/S)",
            ],
            "DIVIDENDS": [
                "Dividend Yield", "Payout Ratio",
            ],
            "RISK": [
                "Beta",
            ],
        }
        
        # Build metrics table
        metric_count = 0
        for section_name, metric_keys in sections.items():
            # Section Header
            section_range = compare_sheet.range(f'A{row}:{chr(65 + len(tickers))}{row}')
            section_range.value = [[section_name] + [""] * len(tickers)]
            section_range.color = (31, 78, 120)
            compare_sheet.range(f'A{row}').api.Font.Color = 0xFFFFFF
            compare_sheet.range(f'A{row}').font.bold = True
            row += 1
            
            # Metrics
            for metric_name in metric_keys:
                compare_sheet.range(f'A{row}').value = metric_name
                
                # Alternating colors
                if metric_count % 2 == 0:
                    row_range = compare_sheet.range(f'A{row}:{chr(65 + len(tickers))}{row}')
                    row_range.color = (242, 242, 242)  # Light gray
                else:
                    row_range = compare_sheet.range(f'A{row}:{chr(65 + len(tickers))}{row}')
                    row_range.color = (255, 255, 255)  # White
                
                # Fill values
                for i, ticker in enumerate(tickers):
                    cell = compare_sheet.range(row, i + 2)
                    cell.value = stock_data[ticker]['metrics'].get(metric_name, 'N/A')
                    cell.api.HorizontalAlignment = -4152  # Right align
                
                row += 1
                metric_count += 1
            
            row += 1  # Blank line
        
        # Format columns
        compare_sheet.range('A:A').column_width = 24
        for i in range(len(tickers)):
            compare_sheet.range(f'{chr(66 + i)}:{chr(66 + i)}').column_width = 18
        
        compare_sheet.range('B7').value = f"Creating chart..."

        # === Comparison Chart (line chart LEFT, summary table RIGHT) ===
        try:
            import matplotlib.dates as mdates

            fig = plt.figure(figsize=(12, 7))
            gs = fig.add_gridspec(1, 2, width_ratios=[3.2, 1.4])

            ax = fig.add_subplot(gs[0, 0])       # chart axis (left)
            ax_tbl = fig.add_subplot(gs[0, 1])   # table axis (right)
            ax_tbl.axis('off')

            colors = ['#2E86DE', '#E67E22', '#27AE60', '#9B59B6', '#E74C3C']
            summary_data = []

            for i, ticker in enumerate(tickers):
                hist = stock_data[ticker]['hist']

                normalized = (hist['Close'] / hist['Close'].iloc[0] - 1) * 100
                ax.plot(
                    normalized.index, normalized,
                    label=ticker,
                    linewidth=2.5,
                    color=colors[i % len(colors)]
                )

                start_price = float(hist['Close'].iloc[0])
                end_price = float(hist['Close'].iloc[-1])
                price_change = end_price - start_price
                pct_change = (price_change / start_price) * 100

                summary_data.append((ticker, start_price, end_price, price_change, pct_change, colors[i % len(colors)]))

            ax.set_title('Stock Comparison (% Change from Start)', fontsize=15, fontweight='bold', pad=12)
            ax.set_xlabel('Date', fontsize=11, fontweight='bold')
            ax.set_ylabel('% Change', fontsize=11, fontweight='bold')
            ax.grid(True, alpha=0.2, linestyle='--')
            ax.axhline(y=0, color='black', linewidth=1, alpha=0.4)

            ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%y'))
            ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')

            ax.legend(loc='upper left', fontsize=10, framealpha=0.95)

            # ----- build summary table on the RIGHT axis -----
            table_rows = []
            for tkr, start, end, chg, pct, c in summary_data:
                sign = "+" if chg >= 0 else ""
                table_rows.append([
                    tkr,
                    f"${start:.2f}",
                    f"${end:.2f}",
                    f"{sign}${chg:.2f}",
                    f"{sign}{pct:.2f}%"
                ])

            tbl = ax_tbl.table(
                cellText=table_rows,
                colLabels=['Ticker', 'Start', 'End', '$ Chg', '% Chg'],
                cellLoc='center',
                colLoc='center',
                loc='center'
            )

            tbl.auto_set_font_size(False)
            tbl.set_fontsize(9)
            tbl.scale(1.05, 1.4)

            # style header row
            for j in range(5):
                cell = tbl[(0, j)]
                cell.set_facecolor('#31508C')
                cell.set_text_props(weight='bold', color='white')

            # style body rows
            for i, (tkr, start, end, chg, pct, c) in enumerate(summary_data):
                row_color = '#F2F2F2' if i % 2 == 0 else 'white'
                for j in range(5):
                    cell = tbl[(i + 1, j)]
                    cell.set_facecolor(row_color)

                    if j == 0:  # ticker column
                        cell.set_text_props(weight='bold', color=c)

                    if j == 4:  # % change column
                        cell.set_text_props(weight='bold', color=('#27AE60' if pct >= 0 else '#E74C3C'))

            plt.tight_layout()

            compare_sheet.pictures.add(
                fig,
                name='ComparisonChart',
                update=True,
                left=compare_sheet.range('F10').left,
                top=compare_sheet.range('F10').top
            )
            plt.close(fig)

        except Exception as e:
            print(f"Error creating comparison chart: {e}")
            compare_sheet.range('B7').value = f"Chart error: {e}"
        
        compare_sheet.range('B7').value = "✓ Comparison complete!"
        
    except Exception as e:
        print(f"Error in compare_stocks: {e}")
        try:
            compare_sheet.range('B7').value = f"Error: {e}"
        except:
            pass

# For testing
if __name__ == '__main__':
    print("Stock Analysis Tool - 3 Sheet Version")
    print("Sheets: Single | News | Compare")
    print("Run from Excel buttons")