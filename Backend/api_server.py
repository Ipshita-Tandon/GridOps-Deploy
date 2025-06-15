from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import json
import tempfile
import shutil
from werkzeug.utils import secure_filename
from xl_extract import extract_cell_content
from xl_json_helper import get_cell_content
from llm_call import analyze_excel_data
from excel_highlighter import highlight_excel_cells
import uuid

app = Flask(__name__)
CORS(app)  # Enable CORS for frontend communication

# Configuration
UPLOAD_FOLDER = 'Excel_files'
EXTRACT_OUTPUT_FOLDER = 'extract-output'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXTRACT_OUTPUT_FOLDER, exist_ok=True)

# Store mapping of file IDs to their data
file_data_cache = {}

# Store mapping of file IDs to their extracted data
file_extracted_data_cache = {}

# Store the original file ID for each uploaded file
original_file_ids = {}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def search_cells_by_value(data, search_term):
    """Search for cells containing specific value"""
    results = []
    search_term = search_term.lower()
    
    for worksheet_name, cells in data.items():
        for cell_coord, cell_data in cells.items():
            if search_term in str(cell_data.get('value', '')).lower():
                results.append({
                    'worksheet': worksheet_name,
                    'coordinate': cell_coord,
                    'value': cell_data.get('value'),
                    'formula': cell_data.get('formula')
                })
    
    return results

@app.route('/upload', methods=['POST'])
def upload_file():
    """
    Upload Excel file endpoint
    Accepts: multipart/form-data with 'file' field
    Returns: JSON with file_id for subsequent queries
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Only .xlsx and .xls files are allowed'}), 400
        
        # Generate unique file ID and secure filename
        file_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{filename}")
        
        # Save the uploaded file
        file.save(file_path)
        
        # Extract cell content
        extract_cell_content(file_path, EXTRACT_OUTPUT_FOLDER)
        
        # Cache the data for quick access
        file_data_cache[file_id] = {
            'filename': filename,
            'file_path': file_path
        }
        
        # Store this as the original file ID
        original_file_ids[file_id] = file_id
        
        print(f"File uploaded successfully. ID: {file_id}, Path: {file_path}")
        print(f"Current cache: {file_data_cache}")
        
        return jsonify({
            'success': True,
            'file_id': file_id,
            'filename': filename,
            'message': 'File uploaded and processed successfully'
        }), 200
        
    except Exception as e:
        print(f"Error in upload_file: {str(e)}")
        return jsonify({'error': f'Error processing file: {str(e)}'}), 500

@app.route('/qna', methods=['POST'])
def ask_question():
    """
    Ask question about Excel content endpoint
    Accepts: JSON with file_id and question
    Returns: JSON with answer based on question type
    """
    try:
        print("Received QnA request")
        data = request.get_json()
        print("Request data:", data)
        
        if not data:
            print("No JSON data provided")
            return jsonify({'error': 'No JSON data provided'}), 400
        
        file_id = data.get('file_id')
        question = data.get('question', '').strip()
        print(f"File ID: {file_id}, Question: {question}")
        print(f"Current cache: {file_data_cache}")
        
        if not file_id or not question:
            print("Missing file_id or question")
            return jsonify({'error': 'file_id and question are required'}), 400
        
        # Check if file exists in cache
        if file_id not in file_data_cache:
            print(f"File {file_id} not found in cache")
            return jsonify({'error': 'File not found. Please upload the file first.'}), 404
        
        print("Loading extracted data...")
        # Load extracted Excel data and analyze with LLM
        if file_id not in file_extracted_data_cache:
            json_path = os.path.join(EXTRACT_OUTPUT_FOLDER, f"{file_id}_{file_data_cache[file_id]['filename'].split('.')[0]}_cell_content.json")
            print(f"Loading JSON from: {json_path}")
            file_extracted_data_cache[file_id] = json.load(open(json_path))
        
        file_extracted_data = file_extracted_data_cache[file_id]
        print("Calling LLM analysis...")
        
        # Analyze data using LLM
        result = analyze_excel_data(file_extracted_data, question)
        print("LLM result:", result)
        
        if result['success']:
            response = {
                'success': True,
                'question': question,
                'answer': result['answer']
            }
            
            if result['error']:  # Add warning if JSON parsing failed
                response['warning'] = result['error']
            
            print("Sending successful response")
            return jsonify(response), 200
        else:
            print("LLM analysis failed")
            return jsonify({
                'error': result['error'],
                'fallback_message': 'LLM analysis failed'
            }), 500

    except Exception as e:
        print("Error in ask_question:", str(e))
        return jsonify({'error': f'Error processing question: {str(e)}'}), 500

@app.route('/highlight', methods=['POST'])
def highlight_excel():
    """
    Highlight specific cells in an Excel file and return the modified file.
    Only two files are kept per upload: the original and the latest highlighted version.
    """
    try:
        data = request.get_json()

        if not data:
            return jsonify({'error': 'No JSON data provided'}), 400

        file_id = data.get('file_id')
        sheet_name = data.get('sheet_name', 'Sheet1')
        cell_ranges = data.get('cell_ranges', 'A1,B2,C3')

        if not file_id:
            return jsonify({'error': 'file_id is required'}), 400

        # Use the original uploaded file
        if file_id not in file_data_cache:
            return jsonify({'error': 'Original file not found. Please upload the file first.'}), 404

        original_info = file_data_cache[file_id]
        original_filename = original_info['filename']
        original_file_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{original_filename}")

        if not os.path.exists(original_file_path):
            return jsonify({'error': 'Original file not found on server'}), 404

        # Always use the same highlighted filename for this file_id
        highlighted_filename = f"{file_id}_highlighted.xlsx"
        highlighted_file_path = os.path.join(UPLOAD_FOLDER, highlighted_filename)

        print(f"Original file path: {original_file_path}")
        print(f"Highlighted file path: {highlighted_file_path}")

        # Only ever delete or overwrite the highlighted file, never the original
        if os.path.exists(highlighted_file_path):
            try:
                os.remove(highlighted_file_path)
            except Exception as e:
                print(f"Warning: Could not remove old highlighted file {highlighted_filename}: {e}")

        # Overwrite the highlighted file every time
        highlight_result = highlight_excel_cells(
            original_file_path,
            sheet_name,
            cell_ranges,
            highlighted_file_path
        )

        if not highlight_result['success']:
            return jsonify({'error': highlight_result['error']}), 500

        # Ensure WOPI public directory exists
        wopi_public_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'Frontend', 'server', 'public')
        os.makedirs(wopi_public_dir, exist_ok=True)

        # Copy new highlighted file to WOPI public dir
        wopi_file_path = os.path.join(wopi_public_dir, highlighted_filename)
        shutil.copy2(highlighted_file_path, wopi_file_path)
        print(f"Copied highlighted file to WOPI directory: {wopi_file_path}")

        # Verify the file exists after copying
        if not os.path.exists(wopi_file_path):
            return jsonify({'error': 'Failed to copy highlighted file to WOPI directory'}), 500

        return jsonify({
            'success': True,
            'file_id': file_id,
            'filename': highlighted_filename,
            'message': 'Highlights applied successfully'
        }), 200

    except Exception as e:
        print(f"Error in highlight_excel: {str(e)}")
        return jsonify({'error': f'Error processing highlight request: {str(e)}'}), 500


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return jsonify({'status': 'healthy', 'service': 'Excel API'}), 200

@app.route('/files', methods=['GET'])
def list_files():
    """List all uploaded files"""
    files_info = []
    for file_id, info in file_data_cache.items():
        files_info.append({
            'file_id': file_id,
            'filename': info['filename'],
            'worksheets': info['worksheets']
        })
    
    return jsonify({
        'success': True,
        'files': files_info
    }), 200

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 