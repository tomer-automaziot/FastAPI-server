"""
Excel Image Extractor API
Extracts embedded images from Excel files and returns them as base64-encoded data.
Designed to work with n8n workflows and Priority ERP integration.
Supports both .xls (legacy) and .xlsx (modern) formats.
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
import zipfile
import xlrd
import struct

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Max file size: 50MB
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


def detect_excel_format(file_bytes):
    """
    Detect if file is XLS or XLSX format.
    
    Args:
        file_bytes: Binary content of the Excel file
        
    Returns:
        'xlsx' or 'xls' or None if unknown
    """
    # XLSX files are ZIP archives starting with PK
    if file_bytes[:4] == b'PK\x03\x04':
        return 'xlsx'
    # XLS files start with Microsoft Compound File signature
    elif file_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        return 'xls'
    return None


def extract_images_from_xlsx(file_bytes):
    """
    Extract all images from an XLSX file using openpyxl.
    """
    images = []
    
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    
    try:
        wb = load_workbook(tmp_path)
        
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            if hasattr(sheet, '_images'):
                for idx, img in enumerate(sheet._images):
                    try:
                        img_data = None
                        img_format = 'png'
                        
                        if hasattr(img, '_data'):
                            img_data = img._data()
                        elif hasattr(img, 'ref'):
                            img_data = img.ref
                            if hasattr(img_data, 'getvalue'):
                                img_data = img_data.getvalue()
                            elif hasattr(img_data, 'read'):
                                img_data = img_data.read()
                        
                        if img_data is None and hasattr(img, '_blob'):
                            img_data = img._blob
                        
                        if img_data:
                            img_format = detect_image_format(img_data)
                            anchor_cell = get_anchor_cell(img)
                            b64_data = base64.b64encode(img_data).decode('utf-8')
                            width, height = get_image_dimensions(img_data)
                            
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
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    
    return images


def extract_images_from_xls(file_bytes):
    """
    Extract all images from an XLS file.
    XLS files store images in the OLE compound document structure.
    """
    images = []
    
    with tempfile.NamedTemporaryFile(suffix='.xls', delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    
    try:
        # Method 1: Try to extract from OLE compound document directly
        images = extract_images_from_ole(file_bytes)
        
        if not images:
            # Method 2: Try using xlrd for basic info, then extract raw images
            logger.info("Trying xlrd method for XLS image extraction")
            images = extract_images_from_xls_raw(file_bytes)
        
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    
    return images


def extract_images_from_ole(file_bytes):
    """
    Extract images from OLE compound document (XLS format).
    Images in XLS are typically stored in the Workbook stream as BIFF records.
    """
    images = []
    
    try:
        # Look for image signatures in the binary data
        # PNG signature
        png_signature = b'\x89PNG\r\n\x1a\n'
        # JPEG signature  
        jpeg_signature = b'\xff\xd8\xff'
        # GIF signatures
        gif_signature1 = b'GIF87a'
        gif_signature2 = b'GIF89a'
        # BMP signature
        bmp_signature = b'BM'
        # EMF signature (Windows Enhanced Metafile)
        emf_signature = b'\x01\x00\x00\x00'
        
        # Find all PNG images
        offset = 0
        while True:
            pos = file_bytes.find(png_signature, offset)
            if pos == -1:
                break
            
            # Try to find the end of PNG (IEND chunk)
            iend_pos = file_bytes.find(b'IEND', pos)
            if iend_pos != -1:
                # PNG IEND chunk is 12 bytes total (4 length + 4 type + 4 CRC)
                img_data = file_bytes[pos:iend_pos + 12]
                if len(img_data) > 100:  # Minimum reasonable size
                    images.append(create_image_entry(img_data, 'png', len(images)))
                    logger.info(f"Found PNG image at offset {pos}, size {len(img_data)}")
            
            offset = pos + 1
        
        # Find all JPEG images
        offset = 0
        while True:
            pos = file_bytes.find(jpeg_signature, offset)
            if pos == -1:
                break
            
            # Find JPEG end marker (FFD9)
            end_marker = b'\xff\xd9'
            end_pos = file_bytes.find(end_marker, pos)
            if end_pos != -1:
                img_data = file_bytes[pos:end_pos + 2]
                if len(img_data) > 100:
                    images.append(create_image_entry(img_data, 'jpeg', len(images)))
                    logger.info(f"Found JPEG image at offset {pos}, size {len(img_data)}")
            
            offset = pos + 1
        
        # Find all GIF images
        for gif_sig in [gif_signature1, gif_signature2]:
            offset = 0
            while True:
                pos = file_bytes.find(gif_sig, offset)
                if pos == -1:
                    break
                
                # GIF ends with trailer byte 0x3B
                end_pos = file_bytes.find(b'\x3b', pos + 6)
                if end_pos != -1:
                    img_data = file_bytes[pos:end_pos + 1]
                    if len(img_data) > 100:
                        images.append(create_image_entry(img_data, 'gif', len(images)))
                        logger.info(f"Found GIF image at offset {pos}, size {len(img_data)}")
                
                offset = pos + 1
        
    except Exception as e:
        logger.error(f"Error in OLE extraction: {str(e)}")
    
    return images


def extract_images_from_xls_raw(file_bytes):
    """
    Alternative method to extract images from XLS by scanning for image headers.
    """
    images = []
    
    # This is a more thorough scan that looks for embedded image data
    # XLS files can have images in various record types
    
    try:
        # Scan for common image format signatures
        signatures = [
            (b'\x89PNG\r\n\x1a\n', 'png'),
            (b'\xff\xd8\xff\xe0', 'jpeg'),
            (b'\xff\xd8\xff\xe1', 'jpeg'),
            (b'\xff\xd8\xff\xdb', 'jpeg'),
            (b'GIF89a', 'gif'),
            (b'GIF87a', 'gif'),
        ]
        
        found_positions = set()  # Avoid duplicates
        
        for signature, fmt in signatures:
            offset = 0
            while True:
                pos = file_bytes.find(signature, offset)
                if pos == -1:
                    break
                
                if pos in found_positions:
                    offset = pos + 1
                    continue
                
                found_positions.add(pos)
                
                # Extract image based on format
                img_data = extract_image_by_format(file_bytes, pos, fmt)
                
                if img_data and len(img_data) > 200:
                    images.append(create_image_entry(img_data, fmt, len(images)))
                    logger.info(f"Extracted {fmt.upper()} from XLS at offset {pos}")
                
                offset = pos + 1
        
    except Exception as e:
        logger.error(f"Error in raw XLS extraction: {str(e)}")
    
    return images


def extract_image_by_format(data, start_pos, fmt):
    """
    Extract image data starting from a given position based on format.
    """
    try:
        if fmt == 'png':
            # PNG ends with IEND chunk followed by CRC
            iend = data.find(b'IEND', start_pos)
            if iend != -1:
                return data[start_pos:iend + 12]
        
        elif fmt == 'jpeg':
            # JPEG ends with FFD9
            end = data.find(b'\xff\xd9', start_pos)
            if end != -1:
                return data[start_pos:end + 2]
        
        elif fmt == 'gif':
            # GIF ends with 0x3B trailer
            # This is simplified - real GIF parsing is more complex
            end = data.find(b'\x3b', start_pos + 10)
            if end != -1:
                return data[start_pos:end + 1]
        
    except Exception as e:
        logger.error(f"Error extracting {fmt}: {str(e)}")
    
    return None


def create_image_entry(img_data, img_format, index):
    """
    Create a standardized image entry dictionary.
    """
    b64_data = base64.b64encode(img_data).decode('utf-8')
    width, height = get_image_dimensions(img_data)
    
    return {
        'index': index,
        'sheet': 'unknown',
        'cell': None,
        'format': img_format,
        'mime_type': f'image/{img_format}',
        'width': width,
        'height': height,
        'size_bytes': len(img_data),
        'base64': b64_data,
        'data_uri': f'data:image/{img_format};base64,{b64_data}'
    }


def detect_image_format(img_data):
    """
    Detect image format from binary data.
    """
    if img_data[:8] == b'\x89PNG\r\n\x1a\n':
        return 'png'
    elif img_data[:2] == b'\xff\xd8':
        return 'jpeg'
    elif img_data[:4] == b'GIF8':
        return 'gif'
    elif img_data[:4] == b'RIFF' and len(img_data) > 12 and img_data[8:12] == b'WEBP':
        return 'webp'
    elif img_data[:2] == b'BM':
        return 'bmp'
    return 'png'


def get_anchor_cell(img):
    """
    Get the anchor cell position for an image.
    """
    anchor_cell = None
    if hasattr(img, 'anchor'):
        if hasattr(img.anchor, '_from'):
            col = img.anchor._from.col
            row = img.anchor._from.row
            anchor_cell = f"{chr(65 + col)}{row + 1}"
    return anchor_cell


def get_image_dimensions(img_data):
    """
    Get image dimensions using PIL.
    """
    width = None
    height = None
    try:
        pil_img = Image.open(io.BytesIO(img_data))
        width, height = pil_img.size
    except:
        pass
    return width, height


def extract_images_from_excel(file_bytes, filename='unknown'):
    """
    Extract all images from an Excel file (supports both XLS and XLSX).
    
    Args:
        file_bytes: Binary content of the Excel file
        filename: Original filename (used for format hint)
        
    Returns:
        List of dictionaries containing image data and metadata
    """
    # Detect format from file content
    file_format = detect_excel_format(file_bytes)
    
    # Fallback to filename extension if detection fails
    if not file_format:
        if filename.lower().endswith('.xlsx'):
            file_format = 'xlsx'
        elif filename.lower().endswith('.xls'):
            file_format = 'xls'
        else:
            # Default to trying xlsx first
            file_format = 'xlsx'
    
    logger.info(f"Detected Excel format: {file_format}")
    
    if file_format == 'xlsx':
        return extract_images_from_xlsx(file_bytes)
    else:
        return extract_images_from_xls(file_bytes)


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint for Railway"""
    return jsonify({
        'status': 'healthy',
        'service': 'excel-image-extractor',
        'supported_formats': ['xls', 'xlsx']
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
        
        # Extract images (now supports both XLS and XLSX)
        images = extract_images_from_excel(file_bytes, filename)
        
        logger.info(f"Successfully extracted {len(images)} images from {filename}")
        
        return jsonify({
            'success': True,
            'filename': filename,
            'format_detected': detect_excel_format(file_bytes),
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