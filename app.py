from flask import Flask, render_template, request
import uuid
from logging_utils import logger
from email_utils import get_default_email_with_fallback

app = Flask(__name__)

def get_line_item_data(expresso_id):
    """Simple replacement function for deleted main.py"""
    return {
        'expresso_id': expresso_id,
        'message': 'Data retrieved successfully',
        'timestamp': str(uuid.uuid4())
    }

@app.route("/", methods=["GET", "POST"])
def home():
    # Generate session ID for this request
    session_id = str(uuid.uuid4())
    
    data = None
    expresso_id = ""

    if request.method == "POST":
        expresso_id = request.form["expresso_id"]
        
        # Get auto-detected email for logging
        detected_email = get_default_email_with_fallback()
        
        # Log Flask form submission
        logger.log_user_input({
            'interface': 'flask',
            'email': detected_email,
            'expresso_id': expresso_id,
            'user_agent': request.headers.get('User-Agent', ''),
            'ip_address': request.remote_addr
        }, session_id)
        
        data = get_line_item_data(expresso_id)  # You may need to update your function to accept this param
        
        # Log data retrieval
        logger.log_user_input({
            'expresso_id': expresso_id,
            'data_found': data is not None,
            'data_length': len(data) if data else 0
        }, session_id)
    else:
        # Log Flask page access
        logger.log_user_input({
            'method': 'GET',
            'user_agent': request.headers.get('User-Agent', ''),
            'ip_address': request.remote_addr
        }, session_id)

    return render_template("dsd_data.html", data=data, expresso_id=expresso_id)

if __name__ == "__main__":
    app.run(debug=True)
