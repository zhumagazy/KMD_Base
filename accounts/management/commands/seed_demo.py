from django.core.management.base import BaseCommand
from django.db import transaction

from accounts.models import User
from exams.models import Assignment, Choice, Question, Test
from learning.models import Course, CourseAccess, CourseGroup, GroupAccess, Material


class Command(BaseCommand):
    help = "Заполняет базу демонстрационными данными (для локальной разработки)."

    @transaction.atomic
    def handle(self, *args, **options):
        admin, created = User.objects.get_or_create(
            username="admin", defaults={"first_name": "Айгерим", "last_name": "Администратор", "is_staff": True, "is_superuser": True}
        )
        if created:
            admin.set_password("admin12345")
            admin.save()
        emp, created = User.objects.get_or_create(
            username="employee", defaults={"first_name": "Алия", "last_name": "Сериккызы", "position": "Менеджер по продажам"}
        )
        if created:
            emp.set_password("employee12345")
            emp.save()
        cand, created = User.objects.get_or_create(
            username="candidate", defaults={"first_name": "Данияр", "last_name": "Ахметов", "is_candidate": True, "position": "Кандидат: дезинфектор"}
        )
        if created:
            cand.set_password("candidate12345")
            cand.save()

        onboarding, _ = CourseGroup.objects.get_or_create(title="Адаптация", defaults={"description": "Всё, что нужно знать в первые недели работы.", "order": 1})
        products, _ = CourseGroup.objects.get_or_create(title="Продукция и услуги", defaults={"order": 2})
        c1, _ = Course.objects.get_or_create(title="Знакомство с компанией", group=onboarding, defaults={"description": "История, миссия, структура и правила.", "order": 1})
        c2, _ = Course.objects.get_or_create(title="Охрана труда и СИЗ", group=onboarding, defaults={"description": "Безопасная работа с дезинфицирующими средствами.", "order": 2})
        c3, _ = Course.objects.get_or_create(title="Дезинфекция: основы", group=products, defaults={"order": 1})
        if not c1.materials.exists():
            Material.objects.create(course=c1, order=1, title="О компании", duration_minutes=5, body="## Кто мы\n\nМы помогаем сохранять здоровье людей.\n\n- Миссия\n- Ценности\n- Структура\n\n> Безопасность — прежде всего.")
            Material.objects.create(course=c1, order=2, kind="video", title="Приветствие руководителя", url="https://www.youtube.com/watch?v=dQw4w9WgXcQ")
            Material.objects.create(course=c1, order=3, kind="link", title="Внутренний портал", url="https://example.com")
        if not c2.materials.exists():
            Material.objects.create(course=c2, order=1, title="Средства индивидуальной защиты", body="Перед началом работ обязательно наденьте **перчатки**, **респиратор** и **очки**.")
        if not c3.materials.exists():
            Material.objects.create(course=c3, order=1, title="Виды дезинфекции", body="Профилактическая, текущая и заключительная.")

        GroupAccess.objects.get_or_create(user=emp, group=onboarding, defaults={"granted_by": admin})
        CourseAccess.objects.get_or_create(user=emp, course=c3, defaults={"granted_by": admin})

        t, created = Test.objects.get_or_create(title="Тест: охрана труда", defaults={"course": c2, "time_limit_minutes": 10, "created_by": admin})
        if created:
            q = Question.objects.create(test=t, order=1, text="Что нужно надеть перед работой с дезсредствами?", kind="multiple")
            Choice.objects.bulk_create([Choice(question=q, text="Перчатки", is_correct=True), Choice(question=q, text="Респиратор", is_correct=True), Choice(question=q, text="Галстук")])
            q = Question.objects.create(test=t, order=2, text="Можно ли смешивать разные дезинфицирующие средства?")
            Choice.objects.bulk_create([Choice(question=q, text="Да"), Choice(question=q, text="Нет", is_correct=True)])
            Question.objects.create(test=t, order=3, kind="text", points=2, text="Опишите порядок действий при попадании средства на кожу.")
        t2, created = Test.objects.get_or_create(title="Входной тест для кандидатов", defaults={"time_limit_minutes": 15, "created_by": admin})
        if created:
            q = Question.objects.create(test=t2, order=1, text="Сколько будет 15% от 200 мл?")
            Choice.objects.bulk_create([Choice(question=q, text="15 мл"), Choice(question=q, text="30 мл", is_correct=True), Choice(question=q, text="20 мл")])
        Assignment.objects.get_or_create(test=t, user=emp, defaults={"created_by": admin})
        Assignment.objects.get_or_create(test=t2, user=cand, defaults={"created_by": admin})
        self.stdout.write(self.style.SUCCESS(
            "Готово. Входы: admin / admin12345, employee / employee12345, candidate / candidate12345"
        ))
