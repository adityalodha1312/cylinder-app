import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from dotenv import load_dotenv
from werkzeug.security import generate_password_hash

# Import models
import sys
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))
from app import app
from db import db
from models import User

def restore():
    with app.app_context():
        # Check if admin exists
        admin_user = User.query.filter_by(username='admin').first()
        if not admin_user:
            admin_user = User(
                username='admin',
                password=generate_password_hash('admin'),
                role='owner',
                name='Admin'
            )
            db.session.add(admin_user)
            db.session.commit()
            print("Successfully restored admin user.")
        else:
            admin_user.password = generate_password_hash('admin')
            db.session.commit()
            print("Successfully reset admin password.")

if __name__ == "__main__":
    restore()
