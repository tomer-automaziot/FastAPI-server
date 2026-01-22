# Excel Image Extractor API

A Python Flask server that extracts embedded images from Excel files and returns them as base64-encoded data. Designed for integration with n8n workflows and Priority ERP.

## Deployment to Railway

### Option 1: Deploy from GitHub

1. Push this folder to a GitHub repository
2. Go to [Railway](https://railway.app)
3. Click "New Project" → "Deploy from GitHub repo"
4. Select your repository
5. Railway will automatically detect the Dockerfile and deploy

### Option 2: Deploy via Railway CLI

```bash
# Install Railway CLI
npm install -g @railway/cli

# Login
railway login

# Create new project
railway init

# Deploy
railway up
```

### After Deployment

1. Go to your Railway project dashboard
2. Click on the service → "Settings" → "Networking"
3. Click "Generate Domain" to get your public URL
4. Your API will be available at: `https://your-service.railway.app`

## API Endpoints

### Health Check
```
GET /health
```
Returns service status.

### Extract Images
```
POST /extract
```

**Input options:**

1. **File Upload (multipart/form-data)**
   ```bash
   curl -X POST -F "file=@spreadsheet.xlsx" https://your-service.railway.app/extract
   ```

2. **Base64 JSON**
   ```json
   {
     "base64": "UEsDBBQAAAAIAA...",
     "filename": "spreadsheet.xlsx"
   }
   ```

3. **Data URI**
   ```json
   {
     "data_uri": "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,UEsDBBQAAAAIAA...",
     "filename": "spreadsheet.xlsx"
   }
   ```

**Response:**
```json
{
  "success": true,
  "filename": "spreadsheet.xlsx",
  "image_count": 3,
  "images": [
    {
      "index": 0,
      "sheet": "Sheet1",
      "cell": "B5",
      "format": "png",
      "mime_type": "image/png",
      "width": 800,
      "height": 600,
      "size_bytes": 45231,
      "base64": "iVBORw0KGgo...",
      "data_uri": "data:image/png;base64,iVBORw0KGgo..."
    }
  ]
}
```

### Extract Simple (Priority-ready)
```
POST /extract-simple
```
Same input as `/extract`, but returns simplified output optimized for Priority ERP:

```json
{
  "success": true,
  "count": 3,
  "data_uris": [
    "data:image/png;base64,iVBORw0KGgo...",
    "data:image/jpeg;base64,/9j/4AAQSkZ..."
  ],
  "images": [
    {"index": 0, "format": "png", "data_uri": "data:image/png;base64,..."}
  ]
}
```

## n8n Integration

### Method 1: Using HTTP Request Node with Binary Data

1. **Get Excel file** (from IMAP, Webhook, etc.)

2. **HTTP Request Node:**
   - Method: POST
   - URL: `https://your-service.railway.app/extract`
   - Body Content Type: `Form-Data/Multipart`
   - Body Parameters:
     - Name: `file`
     - Type: `n8n Binary Data`
     - Input Data Field Name: `attachment_0` (or your binary field name)

3. **Process response** - images are in `$json.images[]`

### Method 2: Using Base64 (if binary doesn't work)

1. **Code Node** to convert binary to base64:
   ```javascript
   const binaryData = await this.helpers.getBinaryDataBuffer(0, 'attachment_0');
   const base64 = binaryData.toString('base64');
   return [{ json: { base64, filename: $binary.attachment_0.fileName } }];
   ```

2. **HTTP Request Node:**
   - Method: POST
   - URL: `https://your-service.railway.app/extract`
   - Body Content Type: `JSON`
   - Body: `{{ $json }}`

### Complete n8n Workflow Example

```
[IMAP Email] 
    ↓
[IF: Has Excel Attachment]
    ↓
[HTTP Request: POST to /extract]
    ↓
[Split In Batches: Process each image]
    ↓
[HTTP Request: Upload to Priority EXTFILES]
```

### n8n Code Node for Priority Upload

```javascript
// After extracting images, upload each to Priority
const images = $input.first().json.images;
const prdiValue = $('Previous Node').first().json.PRDI; // Your PRDI value

const results = [];

for (const img of images) {
  results.push({
    json: {
      EXTFILENAME: `image_${img.index}.${img.format}`,
      EXTFILEDATA: img.data_uri,
      PRDI: prdiValue
    }
  });
}

return results;
```

## Local Development

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run locally
python app.py
```

Server will start at `http://localhost:8080`

## Troubleshooting

### No images extracted
- Ensure the Excel file has embedded images (not linked images)
- Check if images were inserted via "Insert → Pictures → This Device"
- Linked images or images in shapes may not be extracted

### Memory issues
- The server is configured for files up to 50MB
- For larger files, increase `MAX_CONTENT_LENGTH` in app.py

### Timeout errors
- Gunicorn timeout is set to 120 seconds
- For very large files with many images, consider processing in chunks
