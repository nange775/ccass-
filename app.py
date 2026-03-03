from flask import Flask, request, jsonify, render_template
from hkex_scraper import main
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data', methods=['POST'])
def get_data():
    date = request.json.get('date')
    if not date:
        return jsonify({'success': False, 'message': '日期不能为空'})
    
    try:
        success, message, data = main(date)
        return jsonify({'success': success, 'message': message, 'data': data})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

if __name__ == '__main__':
    app.run(debug=True)