import sys
import os

# Add the application directory to the Python path
app_dir = os.path.dirname(os.path.abspath(__file__))
if app_dir not in sys.path:
    sys.path.append(app_dir)

# Import the final v2 Flask application
from final_app import app as application