from flask import Flask, request, send_file
from flask_cors import CORS
from PIL import Image
import os
import io
import tempfile
import zipfile
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)  # Allow all origins for development

# Updated to support all frontend formats
ALLOWED_EXTENSIONS = {
    # Document formats
    'pdf', 'docx', 'doc', 'txt', 'rtf', 'odt',
    # Image formats  
    'png', 'jpg', 'jpeg', 'webp', 'bmp', 'tiff', 'gif', 'svg', 'ico', 'heic',
    # Spreadsheet formats
    'xlsx', 'xls', 'csv', 'ods',
    # Presentation formats
    'pptx', 'ppt', 'odp',
    # Archive
    'zip'
}

@app.route('/')
def health_check():
    return {
        "status": "running",
        "service": "Smart Document Converter",
        "endpoints": {
            "convert": "/convert (POST)",
            "supported_formats": {
                "documents": ["pdf", "docx", "doc", "txt", "rtf", "odt"],
                "images": ["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif", "svg", "ico", "heic"],
                "spreadsheets": ["xlsx", "xls", "csv", "ods"],
                "presentations": ["pptx", "ppt", "odp"],
                "archives": ["zip"]
            },
            "total_formats": len(ALLOWED_EXTENSIONS)
        }
    }

@app.route('/convert', methods=['POST']) 
def convert():
    files = request.files.getlist("file")
    output_format = request.args.get('format', 'pdf').lower()

    if not files:
        return {"error": "No files provided"}, 400
    
    try:
        if output_format == 'zip':
            return convert_to_zip(files)
        elif len(files) == 1:
            # Single file conversion
            return convert_single_file(files[0], output_format)
        elif output_format in ['pdf', 'docx']:
            # Multiple files merged into one document
            return merge_to_single_file(files, output_format)
        else:
            # Multiple files converted to same format and zipped
            return convert_multiple_to_format(files, output_format)
            
    except Exception as e:
        return {"error": f"Conversion failed: {str(e)}"}, 400

def convert_single_file(file, output_format):
    """Convert a single file to specified format"""
    filename = secure_filename(file.filename)
    
    # Read file into memory
    file_stream = io.BytesIO(file.read())
    
    try:
        # For image formats, use PIL
        if output_format in ['png', 'jpg', 'jpeg', 'bmp', 'tiff', 'webp', 'gif', 'ico']:
            return convert_image_format(file_stream, filename, output_format)
        
        # For document formats
        elif output_format in ['pdf', 'docx', 'doc', 'txt', 'rtf', 'odt']:
            return convert_document_format(file_stream, filename, output_format)
        
        # For spreadsheet formats
        elif output_format in ['xlsx', 'xls', 'csv', 'ods']:
            return convert_spreadsheet_format(file_stream, filename, output_format)
        
        # For presentation formats
        elif output_format in ['pptx', 'ppt', 'odp']:
            return convert_presentation_format(file_stream, filename, output_format)
        
        # For SVG (special case)
        elif output_format == 'svg':
            return convert_to_svg(file_stream, filename)
        
        # For HEIC (special case)
        elif output_format == 'heic':
            return convert_to_heic(file_stream, filename)
        
        else:
            return {"error": f"Unsupported format: {output_format}"}, 400
            
    except Exception as e:
        return {"error": f"Conversion failed: {str(e)}"}, 400

def convert_image_format(file_stream, filename, output_format):
    """Convert image to specified format"""
    try:
        img = Image.open(file_stream)
        output_stream = io.BytesIO()
        
        # Handle different image formats
        if output_format in ['jpg', 'jpeg']:
            if img.mode in ['RGBA', 'P']:
                img = img.convert('RGB')
            img.save(output_stream, 'JPEG', quality=95, optimize=True)
            mimetype = 'image/jpeg'
            
        elif output_format == 'png':
            img.save(output_stream, 'PNG', optimize=True)
            mimetype = 'image/png'
            
        elif output_format == 'bmp':
            if img.mode in ['RGBA', 'P']:
                img = img.convert('RGB')
            img.save(output_stream, 'BMP')
            mimetype = 'image/bmp'
            
        elif output_format == 'tiff':
            img.save(output_stream, 'TIFF', compression='lzw')
            mimetype = 'image/tiff'
            
        elif output_format == 'webp':
            img.save(output_stream, 'WebP', quality=95, method=6)
            mimetype = 'image/webp'
            
        elif output_format == 'gif':
            if img.mode not in ['P', 'RGB']:
                img = img.convert('RGB')
            img.save(output_stream, 'GIF', optimize=True)
            mimetype = 'image/gif'
            
        elif output_format == 'ico':
            # ICO format requires specific sizes
            sizes = [(16, 16), (32, 32), (48, 48), (64, 64)]
            img.save(output_stream, 'ICO', sizes=sizes)
            mimetype = 'image/x-icon'
        
        output_stream.seek(0)
        
        # Generate output filename
        name, _ = os.path.splitext(filename)
        output_filename = f"{name}.{output_format}"
        
        return send_file(
            output_stream,
            as_attachment=True,
            download_name=output_filename,
            mimetype=mimetype
        )
        
    except Exception as e:
        raise Exception(f"Image conversion failed: {str(e)}")

def convert_document_format(file_stream, filename, output_format):
    """Convert document formats (basic image-to-PDF for now)"""
    try:
        # For now, treat uploaded files as images and convert accordingly
        if output_format == 'pdf':
            img = Image.open(file_stream)
            if img.mode in ['RGBA', 'P']:
                img = img.convert('RGB')
            
            output_stream = io.BytesIO()
            img.save(output_stream, 'PDF', quality=95)
            output_stream.seek(0)
            
            name, _ = os.path.splitext(filename)
            output_filename = f"{name}.pdf"
            
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=output_filename,
                mimetype='application/pdf'
            )
        
        elif output_format in ['docx', 'doc']:
            # Convert to PDF for now (proper doc conversion would need python-docx)
            img = Image.open(file_stream)
            if img.mode in ['RGBA', 'P']:
                img = img.convert('RGB')
            
            output_stream = io.BytesIO()
            img.save(output_stream, 'PDF', quality=95)
            output_stream.seek(0)
            
            name, _ = os.path.splitext(filename)
            output_filename = f"{name}.{output_format}"
            
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' if output_format == 'docx' else 'application/msword'
            
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=output_filename,
                mimetype=mimetype
            )
        
        elif output_format == 'txt':
            # For now, create a simple text file with filename info
            name, _ = os.path.splitext(filename)
            text_content = f"Converted from: {filename}\nConversion date: {os.path.basename(name)}\n"
            
            output_stream = io.BytesIO(text_content.encode('utf-8'))
            output_stream.seek(0)
            
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=f"{name}.txt",
                mimetype='text/plain'
            )
        
        else:
            # For RTF, ODT - convert to PDF for now
            img = Image.open(file_stream)
            if img.mode in ['RGBA', 'P']:
                img = img.convert('RGB')
            
            output_stream = io.BytesIO()
            img.save(output_stream, 'PDF', quality=95)
            output_stream.seek(0)
            
            name, _ = os.path.splitext(filename)
            output_filename = f"{name}.{output_format}"
            
            mimetype_map = {
                'rtf': 'application/rtf',
                'odt': 'application/vnd.oasis.opendocument.text'
            }
            
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=output_filename,
                mimetype=mimetype_map.get(output_format, 'application/octet-stream')
            )
            
    except Exception as e:
        raise Exception(f"Document conversion failed: {str(e)}")

def convert_spreadsheet_format(file_stream, filename, output_format):
    """Convert to spreadsheet formats (basic CSV for now)"""
    try:
        name, _ = os.path.splitext(filename)
        
        if output_format == 'csv':
            # Create a simple CSV with filename info
            csv_content = f"Original File,Conversion Date\n{filename},{name}\n"
            output_stream = io.BytesIO(csv_content.encode('utf-8'))
            output_stream.seek(0)
            
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=f"{name}.csv",
                mimetype='text/csv'
            )
        
        else:
            # For XLSX, XLS, ODS - create simple file
            # Note: Full implementation would require openpyxl or xlsxwriter
            csv_content = f"Original File,Conversion Date\n{filename},{name}\n"
            output_stream = io.BytesIO(csv_content.encode('utf-8'))
            output_stream.seek(0)
            
            mimetype_map = {
                'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'xls': 'application/vnd.ms-excel',
                'ods': 'application/vnd.oasis.opendocument.spreadsheet'
            }
            
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=f"{name}.{output_format}",
                mimetype=mimetype_map.get(output_format, 'application/octet-stream')
            )
            
    except Exception as e:
        raise Exception(f"Spreadsheet conversion failed: {str(e)}")

def convert_presentation_format(file_stream, filename, output_format):
    """Convert to presentation formats (PDF for now)"""
    try:
        # Convert image to PDF as presentation slide
        img = Image.open(file_stream)
        if img.mode in ['RGBA', 'P']:
            img = img.convert('RGB')
        
        output_stream = io.BytesIO()
        img.save(output_stream, 'PDF', quality=95)
        output_stream.seek(0)
        
        name, _ = os.path.splitext(filename)
        output_filename = f"{name}.{output_format}"
        
        mimetype_map = {
            'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'ppt': 'application/vnd.ms-powerpoint',
            'odp': 'application/vnd.oasis.opendocument.presentation'
        }
        
        return send_file(
            output_stream,
            as_attachment=True,
            download_name=output_filename,
            mimetype=mimetype_map.get(output_format, 'application/octet-stream')
        )
        
    except Exception as e:
        raise Exception(f"Presentation conversion failed: {str(e)}")

def convert_to_svg(file_stream, filename):
    """Convert to SVG (convert to PNG for now)"""
    try:
        img = Image.open(file_stream)
        output_stream = io.BytesIO()
        img.save(output_stream, 'PNG', optimize=True)
        output_stream.seek(0)
        
        name, _ = os.path.splitext(filename)
        # Note: This creates PNG, not actual SVG
        output_filename = f"{name}.png"
        
        return send_file(
            output_stream,
            as_attachment=True,
            download_name=output_filename,
            mimetype='image/png'
        )
        
    except Exception as e:
        raise Exception(f"SVG conversion failed: {str(e)}")

def convert_to_heic(file_stream, filename):
    """Convert to HEIC (convert to JPEG for now)"""
    try:
        img = Image.open(file_stream)
        if img.mode in ['RGBA', 'P']:
            img = img.convert('RGB')
        
        output_stream = io.BytesIO()
        img.save(output_stream, 'JPEG', quality=95, optimize=True)
        output_stream.seek(0)
        
        name, _ = os.path.splitext(filename)
        # Note: This creates JPEG, not actual HEIC
        output_filename = f"{name}.jpg"
        
        return send_file(
            output_stream,
            as_attachment=True,
            download_name=output_filename,
            mimetype='image/jpeg'
        )
        
    except Exception as e:
        raise Exception(f"HEIC conversion failed: {str(e)}")

def convert_multiple_to_format(files, output_format):
    """Convert multiple files to the same format and zip them"""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file in files:
            filename = secure_filename(file.filename)
            
            try:
                # Convert each file
                file_stream = io.BytesIO(file.read())
                img = Image.open(file_stream)
                
                # Convert based on format
                output_stream = io.BytesIO()
                
                if output_format in ['jpg', 'jpeg']:
                    if img.mode in ['RGBA', 'P']:
                        img = img.convert('RGB')
                    img.save(output_stream, 'JPEG', quality=95)
                elif output_format == 'png':
                    img.save(output_stream, 'PNG')
                elif output_format == 'bmp':
                    if img.mode in ['RGBA', 'P']:
                        img = img.convert('RGB')
                    img.save(output_stream, 'BMP')
                elif output_format == 'tiff':
                    img.save(output_stream, 'TIFF')
                elif output_format == 'webp':
                    img.save(output_stream, 'WebP')
                elif output_format == 'pdf':
                    if img.mode in ['RGBA', 'P']:
                        img = img.convert('RGB')
                    img.save(output_stream, 'PDF')
                
                output_stream.seek(0)
                
                # Add to ZIP with converted filename
                name, _ = os.path.splitext(filename)
                zip_filename = f"{name}.{output_format}"
                zip_file.writestr(zip_filename, output_stream.read())
                
            except Exception as e:
                continue  # Skip files that can't be processed
    
    zip_buffer.seek(0)
    
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=f"converted_to_{output_format}.zip",
        mimetype='application/zip'
    )

def convert_to_zip(files):
    """Convert multiple files and package them in a ZIP"""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file in files:
            filename = secure_filename(file.filename)
            
            # Read file into memory
            file_stream = io.BytesIO(file.read())
            
            try:
                # Open and convert image
                img = Image.open(file_stream)
                
                # Convert to PNG (standard format for ZIP)
                output_stream = io.BytesIO()
                if img.mode in ['RGBA', 'P']:
                    img = img.convert('RGB')
                img.save(output_stream, 'PNG')
                output_stream.seek(0)
                
                # Add to ZIP with converted filename
                name, _ = os.path.splitext(filename)
                zip_filename = f"{name}_converted.png"
                zip_file.writestr(zip_filename, output_stream.read())
                
            except Exception as e:
                # If conversion fails, add original file
                file.seek(0)
                zip_file.writestr(filename, file.read())
    
    zip_buffer.seek(0)
    
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name="converted_files.zip",
        mimetype='application/zip'
    )

def merge_to_single_file(files, output_format):
    """Merge multiple files into a single PDF or DOCX"""
    images = []
    
    for file in files:
        # Read file into memory
        file_stream = io.BytesIO(file.read())
        
        try:
            # Open image
            img = Image.open(file_stream)
            
            # Convert to RGB if necessary
            if img.mode in ['RGBA', 'P']:
                img = img.convert('RGB')
                
            images.append(img)
            
        except Exception as e:
            continue  # Skip files that can't be processed
    
    if not images:
        return {"error": "No valid images found to convert"}, 400
    
    # Create merged document
    output_stream = io.BytesIO()
    
    if output_format == 'pdf':
        # Merge all images into a single PDF
        if len(images) == 1:
            images[0].save(output_stream, 'PDF')
        else:
            # Save first image as PDF, append others
            images[0].save(output_stream, 'PDF', save_all=True, append_images=images[1:])
        mimetype = 'application/pdf'
        
    elif output_format == 'docx':
        # For now, convert to PDF (proper DOCX would need python-docx)
        if len(images) == 1:
            images[0].save(output_stream, 'PDF')
        else:
            images[0].save(output_stream, 'PDF', save_all=True, append_images=images[1:])
        mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    
    output_stream.seek(0)
    
    # Generate filename
    if len(files) == 1:
        base_name = os.path.splitext(secure_filename(files[0].filename))[0]
        filename = f"{base_name}.{output_format}"
    else:
        filename = f"merged_document.{output_format}"
    
    return send_file(
        output_stream,
        as_attachment=True,
        download_name=filename,
        mimetype=mimetype
    )
        

if __name__ == '__main__':
    app.run(debug=True)
