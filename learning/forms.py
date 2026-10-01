from django import forms

from .models import Course, CourseGroup, Material


class DateTimeInput(forms.DateTimeInput):
    input_type = "datetime-local"

    def __init__(self, **kwargs):
        super().__init__(format="%Y-%m-%dT%H:%M", **kwargs)


class CourseGroupForm(forms.ModelForm):
    class Meta:
        model = CourseGroup
        fields = ["title", "description", "order"]
        widgets = {"description": forms.Textarea(attrs={"rows": 3})}


class CourseForm(forms.ModelForm):
    class Meta:
        model = Course
        fields = ["title", "group", "description", "is_published", "order"]
        widgets = {"description": forms.Textarea(attrs={"rows": 3})}


class MaterialForm(forms.ModelForm):
    class Meta:
        model = Material
        fields = ["title", "kind", "body", "url", "file", "duration_minutes", "order"]
        widgets = {"body": forms.Textarea(attrs={"rows": 16, "class": "mono"})}


class AccessExpiryForm(forms.Form):
    expires_at = forms.DateTimeField(
        label="Доступ до", required=False, widget=DateTimeInput(),
        help_text="Необязательно. Применяется к вновь отмеченным пользователям.",
    )
