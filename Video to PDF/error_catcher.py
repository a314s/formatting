from flask import Flask, jsonify
import traceback
import sys

app = Flask(__name__)

@app.route('/')
def index():
    return "Error catcher app. Visit /try-import to test imports."

@app.route('/try-import/<module_name>')
def try_import(module_name):
    """Try to import a specific module and return the result"""
    try:
        # Dynamically import the module
        __import__(module_name)
        return jsonify({
            'success': True,
            'message': f'Successfully imported {module_name}'
        })
    except Exception as e:
        # Capture the full traceback
        exc_type, exc_value, exc_traceback = sys.exc_info()
        traceback_details = traceback.format_exception(exc_type, exc_value, exc_traceback)
        
        return jsonify({
            'success': False,
            'error': str(e),
            'traceback': traceback_details
        })

@app.route('/python-info')
def python_info():
    """Return Python information"""
    import sys
    import os
    
    return jsonify({
        'python_version': sys.version,
        'python_path': sys.path,
        'current_dir': os.getcwd(),
        'env_vars': dict(os.environ)
    })

if __name__ == '__main__':
    app.run(debug=True)