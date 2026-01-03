from flask import Flask, render_template
from modules.reconciliation import reconciliation_bp
from modules.invoice_generator import invoice_generator_bp

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB limit

# Register Blueprints
app.register_blueprint(reconciliation_bp, url_prefix='/reconciliation')
app.register_blueprint(invoice_generator_bp, url_prefix='/invoice-generator')

@app.route('/')
def dashboard():
    return render_template('dashboard.html')

if __name__ == '__main__':
    app.run(debug=True, port=1111)
