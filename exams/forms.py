from django import forms
from django.forms import inlineformset_factory

from learning.forms import DateTimeInput

from .models import Assignment, Choice, Question, Test


class TestForm(forms.ModelForm):
    class Meta:
        model = Test
        fields = [
            "title", "description", "course", "time_limit_minutes", "pass_percent", "questions_per_attempt",
            "shuffle_questions", "shuffle_choices", "show_results", "show_correct_answers",
            "require_course_completion", "is_active",
        ]
        widgets = {"description": forms.Textarea(attrs={"rows": 3})}

    def clean_pass_percent(self):
        v = self.cleaned_data["pass_percent"]
        if v > 100:
            raise forms.ValidationError("Не больше 100%.")
        return v


class QuestionForm(forms.ModelForm):
    class Meta:
        model = Question
        fields = ["kind", "text", "image", "points", "explanation", "order"]
        widgets = {"text": forms.Textarea(attrs={"rows": 3}), "explanation": forms.Textarea(attrs={"rows": 2})}


CHOICE_FORMSET_OPTS = dict(fields=["text", "is_correct"], can_delete=True)
ChoiceFormSet = inlineformset_factory(Question, Choice, extra=0, **CHOICE_FORMSET_OPTS)
NewChoiceFormSet = inlineformset_factory(Question, Choice, extra=4, **CHOICE_FORMSET_OPTS)


class AssignForm(forms.Form):
    attempts_allowed = forms.IntegerField(label="Количество попыток", min_value=1, max_value=50, initial=1)
    available_from = forms.DateTimeField(label="Доступен с", required=False, widget=DateTimeInput())
    available_until = forms.DateTimeField(
        label="Доступен до", required=False, widget=DateTimeInput(),
        help_text="После этого времени начать или продолжить тест будет нельзя.",
    )

    def clean(self):
        data = super().clean()
        a, b = data.get("available_from"), data.get("available_until")
        if a and b and b <= a:
            self.add_error("available_until", "Должно быть позже времени начала.")
        return data


class AssignmentEditForm(forms.ModelForm):
    class Meta:
        model = Assignment
        fields = ["attempts_allowed", "available_from", "available_until", "is_revoked"]
        widgets = {"available_from": DateTimeInput(), "available_until": DateTimeInput()}


class GradeForm(forms.Form):
    """Оценка развёрнутых ответов: по полю на каждый ответ."""

    def __init__(self, *args, answers=(), **kwargs):
        super().__init__(*args, **kwargs)
        self.answers = list(answers)
        for a in self.answers:
            self.fields[f"a{a.id}"] = forms.DecimalField(
                label=f"Баллы (из {a.question.points})", min_value=0, max_value=a.question.points,
                decimal_places=2, initial=a.points, required=True,
            )
            a.field_name = f"a{a.id}"
