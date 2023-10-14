from django.db import models

class Datos(models.Model):
    nombre_usuario = models.CharField(max_length=255)
    # Agrega otros campos según tus necesidades
