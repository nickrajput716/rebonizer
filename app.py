"""
REBONIZER v7 – COBRA TECH OFFICIAL EDITION
Team: COBRA TECH (CT)
Colors: #ECF4E8 • #CBF3BB • #ABE7B2 • #93BFC7
Smooth Toggle • Premium UI • Final & Perfect
"""

from flask import Flask, render_template, request, send_file, jsonify
from datetime import datetime
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# Import the scraper
from scraper import SeleniumResultScraper

app = Flask(__name__)
scraper = SeleniumResultScraper()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/scrape', methods=['POST'])
def scrape():
    try:
        data = request.json
        results = scraper.bulk_scrape(data['result_url'], data['start_roll'], data['end_roll'], float(data.get('delay', 3.5)))
        excel = scraper.generate_excel(results)
        if not excel: return jsonify({'error': 'Failed'}), 500
        filename = f"CT_REBONIZER_{data['start_roll']}_to_{data['end_roll']}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        return send_file(excel, as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("="*80)
    print("COBRA TECH REBONIZER v7 LAUNCHED")
    print("Team CT Official Tool → http://127.0.0.1:5000")
    print("="*80)
    app.run(debug=True, port=5000)