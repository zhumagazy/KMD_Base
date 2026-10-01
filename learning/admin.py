from django.contrib import admin

from .models import Course, CourseAccess, CourseGroup, GroupAccess, Material

admin.site.register([CourseGroup, Course, Material, GroupAccess, CourseAccess])
