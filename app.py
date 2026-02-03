from flask import Flask, render_template, request, send_file
import os
import shutil
import tempfile
from process_accounting_report import process_accounting_report

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return "No file part", 400
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    
    if file:
        # Create a temporary directory for this request
        # This ensures thread safety and works on read-only file systems (like Lambda/Render)
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Save uploaded file
            input_path = os.path.join(temp_dir, file.filename)
            file.save(input_path)
            
            # Define output path
            output_filename = f"Processed_{file.filename}"
            output_path = os.path.join(temp_dir, output_filename)
        
            # Process the file
            process_accounting_report(input_path, output_path)
            
            # Send file to user
            # Note: We rely on the OS to eventually clean up /tmp, or we could use cleanup code with after_request
            return send_file(output_path, as_attachment=True)
            
        except Exception as e:
            return f"An error occurred: {str(e)}", 500

if __name__ == '__main__':
    # Easy Deployment Mode with ngrok
    try:
        from pyngrok import ngrok
        # Open a HTTP tunnel on the default port 5000
        # <NgrokTunnel: "http://<public_sub>.ngrok.io" -> "http://localhost:5000">
        public_url = ngrok.connect(5000).public_url
        print(" * ngrok tunnel \"{}\" -> \"http://127.0.0.1:5000\"".format(public_url))
        print(f" * PUBLIC LINK: {public_url}")
        print(" * Share this link with anyone to let them use the app!")
    except Exception as e:
        print(f" * Could not start ngrok: {e}")
        print(" * Running locally only.")

    # Run slightly different port to avoid conflicts
    app.run(debug=True, port=5000)
