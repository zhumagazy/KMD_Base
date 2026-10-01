from django import forms
from django.contrib.auth.forms import SetPasswordForm
from django.contrib.auth.password_validation import validate_password

from .models import User

PROFILE_FIELDS = [
    "last_name", "first_name", "middle_name", "username", "email", "phone", "position",
    "role", "is_candidate", "is_active",
]


class UserCreateForm(forms.ModelForm):
    password1 = forms.CharField(label="Пароль", widget=forms.PasswordInput(render_value=False))
    password2 = forms.CharField(label="Повторите пароль", widget=forms.PasswordInput)

    class Meta:
        model = User
        fields = PROFILE_FIELDS
        labels = {"username": "Логин", "email": "Email", "is_active": "Активен"}
        help_texts = {"username": "Латиница, цифры и символы @.+-_", "is_active": "Неактивные пользователи не могут войти."}

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["first_name"].required = True
        self.fields["last_name"].required = True

    def clean(self):
        data = super().clean()
        p1, p2 = data.get("password1"), data.get("password2")
        if p1 and p2 and p1 != p2:
            self.add_error("password2", "Пароли не совпадают.")
        elif p1:
            try:
                validate_password(p1)
            except forms.ValidationError as e:
                self.add_error("password1", e)
        return data

    def save(self, commit=True):
        user = super().save(commit=False)
        user.set_password(self.cleaned_data["password1"])
        if commit:
            user.save()
        return user


class UserEditForm(forms.ModelForm):
    class Meta:
        model = User
        fields = PROFILE_FIELDS
        labels = UserCreateForm.Meta.labels
        help_texts = UserCreateForm.Meta.help_texts

    def __init__(self, *args, editor=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["first_name"].required = True
        self.fields["last_name"].required = True
        if editor and editor.pk == self.instance.pk:
            # Нельзя лишить себя прав администратора или заблокировать себя.
            self.fields["role"].disabled = True
            self.fields["is_active"].disabled = True


class AdminSetPasswordForm(SetPasswordForm):
    pass
