from django.contrib.auth.models import AbstractUser
from django.db import models


class User(AbstractUser):
    """Пользователь платформы. Ролей всего две: администратор и сотрудник."""

    class Role(models.TextChoices):
        ADMIN = "admin", "Администратор"
        EMPLOYEE = "employee", "Сотрудник"

    role = models.CharField("Роль", max_length=16, choices=Role.choices, default=Role.EMPLOYEE)
    is_candidate = models.BooleanField(
        "Кандидат", default=False, help_text="Отметьте, если это кандидат, а не действующий сотрудник."
    )
    middle_name = models.CharField("Отчество", max_length=150, blank=True)
    position = models.CharField("Должность", max_length=200, blank=True)
    phone = models.CharField("Телефон", max_length=32, blank=True)

    class Meta:
        verbose_name = "Пользователь"
        verbose_name_plural = "Пользователи"
        ordering = ["last_name", "first_name", "username"]

    def save(self, *args, **kwargs):
        if self.is_superuser:
            self.role = self.Role.ADMIN
        super().save(*args, **kwargs)

    @property
    def is_admin(self):
        return self.is_superuser or self.role == self.Role.ADMIN

    def get_full_name(self):
        parts = [self.last_name, self.first_name, self.middle_name]
        return " ".join(p for p in parts if p).strip() or self.username

    def get_short_name(self):
        return self.first_name or self.username

    @property
    def initials(self):
        letters = (self.first_name[:1] + self.last_name[:1]) or self.username[:2]
        return letters.upper()

    @property
    def kind_label(self):
        if self.is_admin:
            return "Администратор"
        return "Кандидат" if self.is_candidate else "Сотрудник"

    def __str__(self):
        return self.get_full_name()
