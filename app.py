"""
Excel Image Extractor API
Extracts embedded images from Excel files and returns them as base64-encoded data.
Designed to work with n8n workflows and Priority ERP integration.
"""

from flask import Flask, request, jsonify
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image
import io
import base64
import tempfile
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Max file size: 50MB
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


def extract_images_from_excel(file_bytes):
    """
    Extract all images from an Excel file.
    
    Args:
        file_bytes: Binary content of the Excel file
        
    Returns:
        List of dictionaries containing image data and metadata
    """
    images = []
    
    # Save to temporary file (openpyxl needs a file path)
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    
    try:
        # Load workbook
        wb = load_workbook(tmp_path)
        
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            # Access images in the sheet
            if hasattr(sheet, '_images'):
                for idx, img in enumerate(sheet._images):
                    try:
                        # Get image data
                        img_data = None
                        img_format = 'png'
                        
                        # Try to get the image blob
                        if hasattr(img, '_data'):
                            img_data = img._data()
                        elif hasattr(img, 'ref'):
                            # For linked images
                            img_data = img.ref
                            if hasattr(img_data, 'getvalue'):
                                img_data = img_data.getvalue()
                            elif hasattr(img_data, 'read'):
                                img_data = img_data.read()
                        
                        if img_data is None:
                            # Alternative: access through _blob
                            if hasattr(img, '_blob'):
                                img_data = img._blob
                        
                        if img_data:
                            # Detect format
                            if img_data[:8] == b'\x89PNG\r\n\x1a\n':
                                img_format = 'png'
                            elif img_data[:2] == b'\xff\xd8':
                                img_format = 'jpeg'
                            elif img_data[:4] == b'GIF8':
                                img_format = 'gif'
                            elif img_data[:4] == b'RIFF' and img_data[8:12] == b'WEBP':
                                img_format = 'webp'
                            
                            # Get position info
                            anchor_cell = None
                            if hasattr(img, 'anchor'):
                                if hasattr(img.anchor, '_from'):
                                    col = img.anchor._from.col
                                    row = img.anchor._from.row
                                    anchor_cell = f"{chr(65 + col)}{row + 1}"
                            
                            # Encode to base64
                            b64_data = base64.b64encode(img_data).decode('utf-8')
                            
                            # Get dimensions if possible
                            width = None
                            height = None
                            try:
                                pil_img = Image.open(io.BytesIO(img_data))
                                width, height = pil_img.size
                            except:
                                pass
                            
                            images.append({
                                'index': len(images),
                                'sheet': sheet_name,
                                'cell': anchor_cell,
                                'format': img_format,
                                'mime_type': f'image/{img_format}',
                                'width': width,
                                'height': height,
                                'size_bytes': len(img_data),
                                'base64': b64_data,
                                'data_uri': f'data:image/{img_format};base64,{b64_data}'
                            })
                            
                            logger.info(f"Extracted image {len(images)} from sheet '{sheet_name}' at cell {anchor_cell}")
                    
                    except Exception as e:
                        logger.error(f"Error extracting image {idx} from sheet {sheet_name}: {str(e)}")
                        continue
        
        wb.close()
        
    finally:
        # Clean up temp file
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    
    return images


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint for Railway"""
    return jsonify({
        'status': 'healthy',
        'service': 'excel-image-extractor'
    })


@app.route('/extract', methods=['POST'])
def extract_images():
    """
    Extract images from an uploaded Excel file.
    
    Accepts:
        - multipart/form-data with 'file' field
        - application/json with 'base64' field containing base64-encoded Excel file
        
    Returns:
        JSON with extracted images as base64 data
    """
    try:
        file_bytes = None
        filename = 'unknown.xlsx'
        
        # Check for file upload
        if 'file' in request.files:
            file = request.files['file']
            filename = file.filename or filename
            file_bytes = file.read()
            logger.info(f"Received file upload: {filename}")
        
        # Check for base64 in JSON body
        elif request.is_json:
            data = request.get_json()
            if 'base64' in data:
                file_bytes = base64.b64decode(data['base64'])
                filename = data.get('filename', filename)
                logger.info(f"Received base64 data for: {filename}")
            elif 'data_uri' in data:
                # Handle data URI format
                data_uri = data['data_uri']
                if ',' in data_uri:
                    file_bytes = base64.b64decode(data_uri.split(',')[1])
                filename = data.get('filename', filename)
                logger.info(f"Received data URI for: {filename}")
        
        # Check for raw binary
        elif request.content_type and 'application/octet-stream' in request.content_type:
            file_bytes = request.data
            logger.info("Received raw binary data")
        
        if not file_bytes:
            return jsonify({
                'success': False,
                'error': 'No file provided. Send file as multipart/form-data, JSON with base64 field, or raw binary.'
            }), 400
        
        # Extract images
        images = extract_images_from_excel(file_bytes)
        
        logger.info(f"Successfully extracted {len(images)} images from {filename}")
        
        return jsonify({
            'success': True,
            'filename': filename,
            'image_count': len(images),
            'images': images
        })
    
    except Exception as e:
        logger.error(f"Error processing request: {str(e)}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/extract-simple', methods=['POST'])
def extract_images_simple():
    """
    Simplified endpoint that returns just the image data URIs.
    Useful for direct integration with Priority ERP.
    
    Returns:
        JSON with array of data URIs ready for Priority attachment upload
    """
    try:
        response = extract_images()
        
        if response[1] != 200 if isinstance(response, tuple) else False:
            return response
        
        data = response.get_json() if hasattr(response, 'get_json') else response
        
        if not data.get('success'):
            return jsonify(data), 400
        
        # Return simplified format
        return jsonify({
            'success': True,
            'count': data['image_count'],
            'data_uris': [img['data_uri'] for img in data['images']],
            'images': [{
                'index': img['index'],
                'format': img['format'],
                'data_uri': img['data_uri']
            } for img in data['images']]
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)
