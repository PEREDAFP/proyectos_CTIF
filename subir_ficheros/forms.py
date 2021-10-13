from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField, IntegerField
from wtforms.validators import DataRequired, Email, Length, Required
from flask_wtf.file import FileField, FileAllowed, FileRequired



class File_Form(FlaskForm):
    num_curso = StringField('Número de curso', validators=[Required("Introduce el número de curso")])
    num_matriculados = IntegerField('Máximo de participantes', validators=[Required("Introduce el número de matriculados")])
    #email = StringField('Email', validators=[DataRequired(), Email('email incorrecto')])
    fecha_inicio = StringField('Fecha Inicio', validators=[Required("Introduce la fecha de inicio"), Length(max=10)])

    fichero = FileField('Imagen de cabecera', validators=[FileRequired(), FileAllowed(['xlsx'],'Ojo con el formato')])

    submit = SubmitField('Procesar')
