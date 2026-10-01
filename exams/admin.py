from django.contrib import admin

from .models import Assignment, Attempt, Choice, Question, Test


class ChoiceInline(admin.TabularInline):
    model = Choice
    extra = 0


@admin.register(Question)
class QuestionAdmin(admin.ModelAdmin):
    inlines = [ChoiceInline]
    list_display = ("text", "test", "kind", "points")


admin.site.register([Test, Assignment, Attempt])
