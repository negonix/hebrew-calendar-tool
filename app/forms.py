from flask_wtf import FlaskForm
from wtforms.fields import RadioField, SubmitField
from flask_wtf.file import FileField, FileRequired

class FileForm(FlaskForm):
    last_years_file = FileField(validators=[FileRequired()])
    this_years_file = FileField(validators=[FileRequired()])
    lang =  RadioField('Label', choices=[('he','Hebrew Dates'),('en','English Dates')])
    submit = SubmitField('Submit')