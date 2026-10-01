from django.contrib import admin
from django.contrib.auth.admin import UserAdmin

from .models import User


@admin.register(User)
class KmdUserAdmin(UserAdmin):
    list_display = ("username", "last_name", "first_name", "role", "is_candidate", "is_active")
    list_filter = ("role", "is_candidate", "is_active")
    fieldsets = UserAdmin.fieldsets + (
        ("Платформа", {"fields": ("role", "is_candidate", "middle_name", "position", "phone")}),
    )
