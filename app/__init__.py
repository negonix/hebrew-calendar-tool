from flask import Flask
import os

app = Flask(__name__)

if app.config["ENV"] == "development":
    app.config.from_object("app.config.DevelopmentConfig")
else:
    app.config.from_object("app.config.ProductionConfig")

from app import views
