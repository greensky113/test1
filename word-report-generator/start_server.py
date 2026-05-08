import os
import io
import zipfile
import base64
from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/')
def index():
    html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '退网报告生成小工具.html')
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    return content, 200, {'Content-Type': 'text/html; charset=utf-8'}

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'message': 'Server running'})

@app.route('/api/generate-reports-batch', methods=['POST'])
def generate_reports_batch():
    try:
        data = request.get_json()
        template_base64 = data.get('template', '')
        data_rows = data.get('dataRows', [])
        community_zips = data.get('communityZips', [])

        if not template_base64:
            return jsonify({'success': False, 'error': '缺少模板文件'}), 400

        template_bytes = base64.b64decode(template_base64)
        community_data_list = []

        for zip_base64 in community_zips:
            zip_bytes = base64.b64decode(zip_base64)
            community_data = parse_community_zip(zip_bytes)
            community_data_list.append(community_data)

        reports = []
        for i, row in enumerate(data_rows):
            community_name = row.get('CommunityName', f'小区{i+1}')
            community_data = community_data_list[i] if i < len(community_data_list) else (community_data_list[0] if community_data_list else {})
            
            try:
                report_bytes = template_bytes
                report_base64 = base64.b64encode(report_bytes).decode('utf-8')
                reports.append({
                    'index': i,
                    'communityName': community_name,
                    'report': report_base64
                })
            except Exception as e:
                reports.append({
                    'index': i,
                    'communityName': community_name,
                    'error': str(e)
                })

        return jsonify({
            'success': True,
            'reports': reports
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

def parse_community_zip(zip_bytes):
    result = {
        'excel_files': {},
        'environment_images': {},
        'removed_equipment_images': {}
    }

    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                name = name.replace('\\', '/')

                if name.startswith('excel/') and name.endswith('.xlsx'):
                    filename = os.path.basename(name)
                    with zf.open(name) as f:
                        result['excel_files'][filename] = base64.b64encode(f.read()).decode('utf-8')

                elif name.startswith('Environment_images/') and (name.endswith('.jpg') or name.endswith('.png')):
                    filename = os.path.basename(name)
                    with zf.open(name) as f:
                        result['environment_images'][filename] = base64.b64encode(f.read()).decode('utf-8')

                elif name.startswith('RemovedEquipment_images/') and (name.endswith('.jpg') or name.endswith('.png')):
                    filename = os.path.basename(name)
                    with zf.open(name) as f:
                        result['removed_equipment_images'][filename] = base64.b64encode(f.read()).decode('utf-8')

    except Exception as e:
        print(f"Error parsing zip: {e}")

    return result

if __name__ == '__main__':
    port = 8080
    print(f"Server running on http://localhost:{port}")
    app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)