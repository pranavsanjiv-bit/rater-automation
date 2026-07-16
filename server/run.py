import sys
import os

# Add the parent directory to the path so we can import from server package
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from server import create_app

app = create_app()

if __name__ == "__main__":
    app.run(port=5050, debug=True)