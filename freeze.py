from flask_frozen import Freezer
from app import app  # your Flask app

freezer = Freezer(app)

if __name__ == '__main__':
    freezer.freeze()