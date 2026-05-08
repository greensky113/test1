# -*- coding: utf-8 -*-
import os
import io
import re
import json
import base64
import zipfile
import tempfile
import shutil
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

class OleEmbedder:
    def __init__(self):
        self.rel_id_counter = 100
    
    def generate_rel_id(self):
        rel_id = f"rId{self.rel_id_counter}"
        self.rel_id_counter += 1
        return rel_id
    
    def create_ole_object_xml(self, rel_id, file_name):
        shape_id = rel_id.replace('rId', '')
        ole_xml = f'''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:r>
    <w:object w:dxaOrig="21600" w:dyaOrig="21600">
      <v:shape id="_x0000_s{shape_id}" type="#_x0000_t75" style="width:192pt;height:144pt" o:ole="true">
        <v:imagedata r:id="{rel_id}" o:title="{file_name}"/>
        <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_s{shape_id}" DrawAspect="Content" ObjectID="_x0000_o{shape_id}"/>
      </v:shape>
    </w:object>
  </w:r>
</w:p>'''
        return ole_xml
    
    def create_ole_relationship_xml(self, rel_id):
        return f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/Microsoft_Excel_Worksheet.xlsx" TargetMode="Internal"/>'
    
    def embed_excel_to_docx(self, docx_bytes, excel_files, placeholder_mappings):
        """
        将Excel文件作为OLE对象嵌入到docx模板中
        
        Args:
            docx_bytes: docx模板文件的字节数据
            excel_files: dict, {占位符名: excel文件bytes}
            placeholder_mappings: dict, {占位符名: Excel文件名}
        
        Returns:
            修改后的docx文件字节数据
        """
        with tempfile.TemporaryDirectory() as temp_dir:
            input_zip = zipfile.ZipFile(io.BytesIO(docx_bytes), 'r')
            input_zip.extractall(temp_dir)
            input_zip.close()
            
            document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
            rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
            
            with open(document_xml_path, 'r', encoding='utf-8') as f:
                document_xml = f.read()
            
            with open(rels_path, 'r', encoding='utf-8') as f:
                rels_xml = f.read()
            
            embeddings_dir = os.path.join(temp_dir, 'word', 'embeddings')
            os.makedirs(embeddings_dir, exist_ok=True)
            
            for placeholder, excel_data in excel_files.items():
                excel_file_name = placeholder_mappings.get(placeholder, f'{placeholder}.xlsx')
                embed_file_name = f'Microsoft_Excel_Worksheet_{self.rel_id_counter}.xlsx'
                
                excel_path = os.path.join(embeddings_dir, embed_file_name)
                with open(excel_path, 'wb') as f:
                    f.write(excel_data)
                
                rel_id = self.generate_rel_id()
                ole_rel_xml = self.create_ole_relationship_xml(rel_id)
                ole_object_xml = self.create_ole_object_xml(rel_id, excel_file_name)
                
                document_xml = document_xml.replace(placeholder, ole_object_xml)
                rels_xml = rels_xml.replace('</Relationships>', f'{ole_rel_xml}</Relationships>')
                
                app.logger.info(f'已嵌入Excel文件: {excel_file_name} 到占位符 {placeholder}')
            
            with open(document_xml_path, 'w', encoding='utf-8') as f:
                f.write(document_xml)
            
            with open(rels_path, 'w', encoding='utf-8') as f:
                f.write(rels_xml)
            
            output_buffer = io.BytesIO()
            with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as output_zip:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, temp_dir)
                        output_zip.write(file_path, arcname)
            
            return output_buffer.getvalue()


@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'message': 'Word Report Generator API'})


@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    try:
        data = request.get_json()
        
        template_base64 = data.get('template')
        excel_files_data = data.get('excelFiles', [])
        placeholder_mappings = data.get('placeholderMappings', {})
        
        if not template_base64:
            return jsonify({'error': '缺少模板文件'}), 400
        
        template_bytes = base64.b64decode(template_base64)
        
        excel_files = {}
        for item in excel_files_data:
            name = item.get('name')
            data_bytes = base64.b64decode(item.get('data'))
            excel_files[name] = data_bytes
        
        embedder = OleEmbedder()
        result_bytes = embedder.embed_excel_to_docx(
            template_bytes, 
            excel_files, 
            placeholder_mappings
        )
        
        result_base64 = base64.b64encode(result_bytes).decode('utf-8')
        
        return jsonify({
            'success': True,
            'report': result_base64
        })
    
    except Exception as e:
        app.logger.error(f'生成报告失败: {str(e)}')
        return jsonify({'error': str(e)}), 500


@app.route('/api/extract-placeholders', methods=['POST'])
def extract_placeholders():
    try:
        if 'template' not in request.files:
            return jsonify({'error': '缺少模板文件'}), 400
        
        template_file = request.files['template']
        docx_bytes = template_file.read()
        
        with tempfile.TemporaryDirectory() as temp_dir:
            input_zip = zipfile.ZipFile(io.BytesIO(docx_bytes), 'r')
            input_zip.extractall(temp_dir)
            input_zip.close()
            
            document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
            
            with open(document_xml_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            excel_placeholders = re.findall(r'\{\{excel:([^}]+)\}\}', content)
            text_placeholders = re.findall(r'\{\{([^@#][^}]+)\}\}', content)
            text_placeholders = [p for p in text_placeholders if not any(p.startswith(prefix) for prefix in ['excel:', '@', '#'])]
            
            return jsonify({
                'success': True,
                'placeholders': {
                    'excel': excel_placeholders,
                    'text': text_placeholders
                }
            })
    
    except Exception as e:
        app.logger.error(f'提取占位符失败: {str(e)}')
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
