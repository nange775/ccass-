from flask import Flask, request, jsonify, render_template
from hkex_scraper import main
from hkex_all import main as batch_main
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data', methods=['POST'])
def get_data():
    date = request.json.get('date')
    file_path = request.json.get('file_path', '.')
    stock_code = request.json.get('stock_code', '02158')
    if not date:
        return jsonify({'success': False, 'message': '日期不能为空'})
    
    try:
        success, message, data = main(query_date=date, file_path=file_path, stock_code=stock_code)
        return jsonify({'success': success, 'message': message, 'data': data})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/batch', methods=['POST'])
def get_batch_data():
    start_date = request.json.get('start_date')
    end_date = request.json.get('end_date')
    file_path = request.json.get('file_path', '.')
    stock_code = request.json.get('stock_code', '02158')
    
    if not start_date or not end_date:
        return jsonify({'success': False, 'message': '开始日期和结束日期不能为空'})
    
    try:
        # 调用hkex_all.py中的批量查询功能，传递参数
        success, message, data = batch_main(
            start_date=start_date,
            end_date=end_date,
            stock_code=stock_code,
            file_path=file_path
        )
        return jsonify({'success': success, 'message': message, 'data': data})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/batch')
def batch_page():
    return render_template('batch.html')

if __name__ == '__main__':
    app.run(debug=True)
