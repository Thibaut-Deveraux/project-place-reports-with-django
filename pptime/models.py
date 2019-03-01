from django.db import models
import os
import PiWi.settings

class PPTimeReport(models.Model):
    docfile = models.FileField(upload_to='timereports/')
    creationdate = models.DateTimeField(auto_now_add=True)
    comment = models.CharField(max_length=250, blank=True)
