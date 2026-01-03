from flask import Blueprint

create_invoice_bp = Blueprint('create_invoice', __name__, 
                              template_folder='../../templates/create_invoice',
                              static_folder='../../static/create_invoice')

from . import routes
