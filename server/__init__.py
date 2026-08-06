# pyrefly: ignore [missing-import]
from flask import Flask
from server.db import init_db
from server.routes.templates import templates_bp
from server.routes.data_process import processing_bp
from server.routes.injestion import injestion_bp
from server.routes.dimension_library import dimension_library_bp

def create_app():
    """
    Create and configure the Flask application.

    Initializes the database tables and registers all application blueprints.
    """
    app = Flask(__name__)

    # Initialize DB tables on startup
    init_db()

    # Register application blueprints
    app.register_blueprint(templates_bp)
    app.register_blueprint(processing_bp)
    app.register_blueprint(injestion_bp)
    app.register_blueprint(dimension_library_bp)

    return app