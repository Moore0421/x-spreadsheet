from flask import Flask, request, jsonify
import os
import json
from werkzeug.utils import secure_filename

app = Flask(__name__)

# 设置上传文件大小限制
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB

# 设置数据目录
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")

# 确保数据目录存在
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# 获取文件名函数
def get_file_name(id, pure):
    file_name = f"spreadsheet-data-pure-{id}.json" if pure else f"spreadsheet-data-{id}.json"
    return os.path.join(DATA_DIR, file_name)

# 设置CORS
@app.after_request
def add_cors_headers(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type')
    return response

# 分片上传接口
@app.route('/api/chunk-upload', methods=['POST'])
def chunk_upload():
    chunk_index = request.args.get('chunkIndex')
    total_chunks = request.args.get('totalChunks')
    id = request.args.get('id')
    data = json.dumps(request.json)
    temp_dir = os.path.join(os.path.dirname(__file__), "temp")
    
    # 确保临时目录存在
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    
    chunk_path = os.path.join(temp_dir, f"{id}-{chunk_index}")
    
    try:
        with open(chunk_path, 'w') as f:
            f.write(data)
        return jsonify({
            'success': True,
            'chunkIndex': chunk_index,
            'message': f"Chunk {chunk_index} of {total_chunks} received"
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# 合并分片接口
@app.route('/api/merge-chunks', methods=['POST'])
def merge_chunks():
    id = request.args.get('id')
    total_chunks = int(request.args.get('totalChunks'))
    pure = request.args.get('pure') == 'true'
    temp_dir = os.path.join(os.path.dirname(__file__), "temp")
    final_path = get_file_name(id, pure)
    raw_data = ""
    
    try:
        print(f"开始合并，总分片数: {total_chunks}")
        
        # 按顺序读取并合并所有分片
        for i in range(total_chunks):
            chunk_path = os.path.join(temp_dir, f"{id}-{i}")
            with open(chunk_path, 'r') as f:
                chunk_content = f.read()
            chunk_data = json.loads(chunk_content)
            
            # 提取分片中的实际数据
            if chunk_data and 'data' in chunk_data:
                raw_data += chunk_data['data']
            
            # 删除分片文件
            os.remove(chunk_path)
        
        # 解析完整的JSON字符串
        final_data = json.loads(raw_data)
        print(f"合并后数据对象类型: {type(final_data)}")
        
        # 保存为格式化的JSON
        with open(final_path, 'w') as f:
            json.dump(final_data, f, indent=2)
        print("最终文件已保存")
        
        return jsonify({'success': True, 'message': 'File merged successfully'})
    except Exception as e:
        print(f"合并错误: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

# 获取表格总数接口
@app.route('/api/getSheetCount', methods=['GET'])
def get_sheet_count():
    id = request.args.get('id')
    file_name = get_file_name(id, False)
    print(f"fileName {file_name}")
    
    try:
        if os.path.exists(file_name):
            with open(file_name, 'r') as f:
                data = json.load(f)
            return jsonify({
                'success': True,
                'count': len(data)
            })
        else:
            return jsonify({'success': False})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# 获取单个表格数据接口
@app.route('/api/getSheet', methods=['GET'])
def get_sheet():
    id = request.args.get('id')
    index = int(request.args.get('index'))
    file_name = get_file_name(id, False)
    
    try:
        if os.path.exists(file_name):
            with open(file_name, 'r') as f:
                data = json.load(f)
            if 0 <= index < len(data):
                return jsonify({
                    'success': True,
                    'sheet': data[index]
                })
            else:
                return jsonify({'success': False, 'message': 'Sheet index out of range'})
        else:
            return jsonify({})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

# 深度比较两个对象是否相同
def is_equal(obj1, obj2):
    return json.dumps(obj1) == json.dumps(obj2)

# 递归更新对象
def update_object(full_obj, pure_obj):
    is_modified = False
    
    # 特殊处理 rows 对象
    if 'rows' in pure_obj and 'rows' in full_obj:
        for row_key in pure_obj['rows']:
            if row_key in full_obj['rows']:
                # 处理 cells 对象
                if 'cells' in pure_obj['rows'][row_key] and 'cells' in full_obj['rows'][row_key]:
                    for cell_key in pure_obj['rows'][row_key]['cells']:
                        if cell_key in full_obj['rows'][row_key]['cells']:
                            # 检查并更新 text 字段
                            if ('text' in pure_obj['rows'][row_key]['cells'][cell_key] and
                                not is_equal(full_obj['rows'][row_key]['cells'][cell_key]['text'],
                                            pure_obj['rows'][row_key]['cells'][cell_key]['text'])):
                                full_obj['rows'][row_key]['cells'][cell_key]['text'] = pure_obj['rows'][row_key]['cells'][cell_key]['text']
                                full_obj['rows'][row_key]['cells'][cell_key]['formattedText'] = pure_obj['rows'][row_key]['cells'][cell_key]['text']
                                is_modified = True
    
    # 处理其他普通字段
    for key in pure_obj:
        if key == 'rows':
            continue  # 跳过已处理的 rows
        
        if key in full_obj:
            if (isinstance(pure_obj[key], dict) and pure_obj[key] is not None and
                isinstance(full_obj[key], dict) and full_obj[key] is not None):
                modified, updated_obj = update_object(full_obj[key], pure_obj[key])
                if modified:
                    full_obj[key] = updated_obj
                    is_modified = True
            elif not is_equal(full_obj[key], pure_obj[key]):
                full_obj[key] = pure_obj[key]
                is_modified = True
    
    return is_modified, full_obj

# 同步纯数据到完整文件接口
@app.route('/api/sync-data', methods=['POST'])
def sync_data():
    id = request.args.get('id')
    full_file_name = get_file_name(id, False)
    
    # 检查完整数据文件是否存在
    if not os.path.exists(full_file_name):
        return jsonify({
            'success': False,
            'message': 'Full data file not found'
        }), 404
    
    # 处理上传的文件
    if 'pureDataFile' not in request.files:
        return jsonify({'success': False, 'message': 'No file part'}), 400
    
    file = request.files['pureDataFile']
    if file.filename == '':
        return jsonify({'success': False, 'message': 'No selected file'}), 400
    
    try:
        # 创建临时目录
        uploads_dir = os.path.join(os.path.dirname(__file__), "uploads")
        if not os.path.exists(uploads_dir):
            os.makedirs(uploads_dir)
        
        # 保存上传的文件
        file_path = os.path.join(uploads_dir, secure_filename(file.filename))
        file.save(file_path)
        
        # 读取上传的纯数据文件
        with open(file_path, 'r') as f:
            pure_data = json.load(f)
        
        # 读取完整数据文件
        with open(full_file_name, 'r') as f:
            full_data = json.load(f)
        
        modified_sheets = 0
        
        # 遍历纯数据文件中的每个对象
        for pure_sheet in pure_data:
            # 在完整数据中找到对应的对象
            full_sheet = next((sheet for sheet in full_data if sheet['name'] == pure_sheet['name']), None)
            
            if full_sheet:
                # 递归更新对象
                modified, updated_obj = update_object(full_sheet, pure_sheet)
                if modified:
                    full_sheet.update(updated_obj)
                    modified_sheets += 1
        
        # 删除临时上传的文件
        os.remove(file_path)
        
        # 只有在有修改时才保存文件
        if modified_sheets > 0:
            with open(full_file_name, 'w') as f:
                json.dump(full_data, f, indent=2)
            return jsonify({
                'success': True,
                'message': f'Successfully synchronized data. Modified {modified_sheets} sheets.'
            })
        else:
            return jsonify({
                'success': True,
                'message': 'No changes needed, data already in sync.'
            })
    except Exception as e:
        print(f"同步数据错误: {e}")
        # 确保清理临时文件
        if os.path.exists(file_path):
            os.remove(file_path)
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=3000)