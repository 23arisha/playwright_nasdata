from flask import Flask, send_file, jsonify
from nasdata import run_scraper
import os

app = Flask(__name__)

@app.route("/")
def index():
    return '''
    <h2>Stockhouse NASDAQ Scraper</h2>
    <a href="/download">Download Excel Report</a>
    '''

@app.route('/download')
def download():
    filepath = run_scraper()
    return send_file(filepath, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
