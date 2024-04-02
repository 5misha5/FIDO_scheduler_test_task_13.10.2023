from django.db import models
from django.core.validators import MinValueValidator, MaxValueValidator


class Courses(models.Model):
    course_id = models.AutoField(primary_key=True)
    course_name = models.CharField(max_length=255, null=False)
    lector_name = models.CharField(max_length=255, null=True)
    group_number = models.CharField(max_length=255, null=False)
    day_of_week = models.IntegerField(null=False, validators=[MinValueValidator(0), MaxValueValidator(6)])
    weeks = models.ArrayField(models.IntegerField(), null=False)
    lesson_number = models.IntegerField(null=False)
    auditory_name = models.CharField(max_length=255, null=False)
