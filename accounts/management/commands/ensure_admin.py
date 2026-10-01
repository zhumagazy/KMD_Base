import os

from django.core.management.base import BaseCommand

from accounts.models import User


class Command(BaseCommand):
    help = "Создаёт администратора из переменных ADMIN_USERNAME / ADMIN_PASSWORD, если его ещё нет."

    def handle(self, *args, **options):
        username = os.environ.get("ADMIN_USERNAME")
        password = os.environ.get("ADMIN_PASSWORD")
        if not username or not password:
            self.stdout.write("ADMIN_USERNAME/ADMIN_PASSWORD не заданы — пропускаю.")
            return
        if User.objects.filter(username=username).exists():
            self.stdout.write(f"Пользователь {username} уже существует.")
            return
        User.objects.create_superuser(
            username=username, password=password, email=os.environ.get("ADMIN_EMAIL", ""),
            first_name="Администратор",
        )
        self.stdout.write(self.style.SUCCESS(f"Администратор {username} создан."))
