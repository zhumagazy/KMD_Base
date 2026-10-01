from pathlib import Path

from django.conf import settings
from django.core.management.base import BaseCommand, CommandError

from learning.importer import Importer


class Command(BaseCommand):
    help = "Импортирует группы, курсы, материалы и тесты из папки (по умолчанию content/)."

    def add_arguments(self, parser):
        parser.add_argument("path", nargs="?", default=str(Path(settings.BASE_DIR) / "content"))
        parser.add_argument("--dry-run", action="store_true", help="Только показать, что будет сделано.")
        parser.add_argument("--quiet", action="store_true", help="Не выводить список, только итог.")

    def handle(self, path, dry_run, quiet, **options):
        if not Path(path).is_dir():
            self.stdout.write(f"Папка {path} не найдена — импорт пропущен.")
            return
        try:
            report = Importer(path, dry_run=dry_run).run()
        except FileNotFoundError as e:
            raise CommandError(str(e))
        if not quiet:
            for line in report.created:
                self.stdout.write(self.style.SUCCESS(f"+ {line}"))
            for line in report.updated:
                self.stdout.write(f"~ {line}")
        for line in report.skipped:
            self.stdout.write(self.style.WARNING(f"! {line}"))
        for line in report.errors:
            self.stdout.write(self.style.ERROR(f"✗ {line}"))
        prefix = "[проверка, ничего не сохранено] " if dry_run else ""
        self.stdout.write(self.style.SUCCESS(
            f"{prefix}Создано: {len(report.created)}, обновлено: {len(report.updated)}, "
            f"без изменений: {report.unchanged}, пропущено: {len(report.skipped)}, ошибок: {len(report.errors)}"
        ))
