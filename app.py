from flask import Flask, render_template_string, request, send_file
import os
import shutil
import tempfile
from process_accounting_report import process_accounting_report

app = Flask(__name__)

# Embedded HTML to avoid "TemplateNotFound" errors on cloud platforms
INDEX_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Accounting Report Processor</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600&display=swap');

        :root {
            --primary-bg: #0f172a;
            --glass-bg: rgba(255, 255, 255, 0.05);
            --glass-border: rgba(255, 255, 255, 0.1);
            --accent-color: #3b82f6;
            --accent-glow: rgba(59, 130, 246, 0.5);
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
        }

        body {
            margin: 0;
            padding: 0;
            font-family: 'Outfit', sans-serif;
            background-color: var(--primary-bg);
            background-image: 
                radial-gradient(circle at 10% 20%, rgba(59, 130, 246, 0.15) 0%, transparent 40%),
                radial-gradient(circle at 90% 80%, rgba(236, 72, 153, 0.15) 0%, transparent 40%);
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            color: var(--text-main);
            overflow: hidden;
        }

        .container {
            background: var(--glass-bg);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid var(--glass-border);
            border-radius: 24px;
            padding: 3rem;
            width: 100%;
            max-width: 500px;
            text-align: center;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
            animation: fadeIn 0.8s ease-out;
        }

        h1 {
            font-weight: 600;
            margin-bottom: 0.5rem;
            background: linear-gradient(to right, #60a5fa, #e879f9);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
            font-size: 2.2rem;
        }

        p {
            color: var(--text-muted);
            margin-bottom: 2.5rem;
            line-height: 1.6;
        }

        .upload-area {
            border: 2px dashed var(--glass-border);
            border-radius: 16px;
            padding: 2.5rem 1.5rem;
            margin-bottom: 2rem;
            transition: all 0.3s ease;
            cursor: pointer;
            position: relative;
            background: rgba(0,0,0,0.2);
        }

        .upload-area:hover, .upload-area.dragover {
            border-color: var(--accent-color);
            background: rgba(59, 130, 246, 0.1);
            box-shadow: 0 0 20px var(--accent-glow);
        }

        .upload-icon {
            font-size: 3rem;
            color: var(--text-muted);
            margin-bottom: 1rem;
            display: block;
        }

        input[type="file"] {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            opacity: 0;
            cursor: pointer;
        }

        .file-name {
            display: block;
            margin-top: 1rem;
            font-size: 0.9rem;
            color: var(--accent-color);
            font-weight: 600;
            min-height: 1.2em;
        }

        button {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
            border: none;
            padding: 1rem 2.5rem;
            color: white;
            font-size: 1rem;
            font-weight: 600;
            border-radius: 12px;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(59, 130, 246, 0.4);
            width: 100%;
        }

        button:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 25px rgba(59, 130, 246, 0.5);
        }

        button:active {
            transform: translateY(0);
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .loading {
            display: none;
            margin-top: 1rem;
            color: var(--text-muted);
            font-size: 0.9rem;
            animation: pulse 1.5s infinite;
        }

        @keyframes pulse {
            0%, 100% { opacity: 0.6; }
            50% { opacity: 1; }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Accounting Processor</h1>
        <p>Drop your report to automatically process journal entries and errors.</p>
        
        <form action="/process" method="post" enctype="multipart/form-data" id="upload-form">
            <div class="upload-area" id="drop-zone">
                <span class="upload-icon">📂</span>
                <span>Drag & drop or Click to Browse</span>
                <span class="file-name" id="file-name"></span>
                <input type="file" name="file" id="file-input" accept=".xlsx" required>
            </div>
            
            <button type="submit" id="submit-btn">Process Report</button>
            <div class="loading" id="loading-msg">Processing your file... please wait</div>
        </form>
    </div>

    <script>
        const fileInput = document.getElementById('file-input');
        const fileNameDisplay = document.getElementById('file-name');
        const dropZone = document.getElementById('drop-zone');
        const form = document.getElementById('upload-form');
        const submitBtn = document.getElementById('submit-btn');
        const loadingMsg = document.getElementById('loading-msg');

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                fileNameDisplay.textContent = e.target.files[0].name;
                dropZone.style.borderColor = '#3b82f6';
                dropZone.style.background = 'rgba(59, 130, 246, 0.1)';
            }
        });

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            fileInput.files = e.dataTransfer.files;
            if (fileInput.files.length > 0) {
                fileNameDisplay.textContent = fileInput.files[0].name;
                fileInput.dispatchEvent(new Event('change'));
            }
        });

        form.addEventListener('submit', () => {
            submitBtn.style.opacity = '0.7';
            submitBtn.textContent = 'Processing...';
            submitBtn.disabled = true;
            loadingMsg.style.display = 'block';
        });
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(INDEX_HTML)

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return "No file part", 400
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    
    if file:
        temp_dir = tempfile.mkdtemp()
        try:
            input_path = os.path.join(temp_dir, file.filename)
            file.save(input_path)
            
            output_filename = f"Processed_{file.filename}"
            output_path = os.path.join(temp_dir, output_filename)
        
            process_accounting_report(input_path, output_path)
            
            return send_file(output_path, as_attachment=True)
            
        except Exception as e:
            return f"An error occurred: {str(e)}", 500

if __name__ == '__main__':
    # Only try ngrok if NOT on Render
    if not os.environ.get('RENDER'):
        try:
            from pyngrok import ngrok
            public_url = ngrok.connect(5000).public_url
            print(f" * PUBLIC LINK: {public_url}")
        except Exception as e:
            print(f" * ngrok skipped: {e}")

    app.run(debug=True, port=int(os.environ.get('PORT', 5000)))
