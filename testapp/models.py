from django.db import models


# Create your models here.
class TestModel(models.Model):
    text = models.CharField(max_length=255)
    number = models.IntegerField()
    timestamp = models.DateTimeField(auto_now_add=True)
