from django.db import models

# Create your models here.
class File_Upload(models.Model):
    file_name=models.CharField(max_length=100)
    file=models.FileField()
    