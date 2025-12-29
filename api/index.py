from app import app

# Vercel entry point
# No need for handler(), Vercel/WSGI handles 'app' object automatically if exposed.
# But for @vercel/python, we usually expose 'app' as a variable.
