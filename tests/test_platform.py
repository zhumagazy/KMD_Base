from datetime import timedelta

from django.core.management import call_command
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from accounts.models import User
from exams.models import Assignment, Attempt, Question, Test
from learning.models import Course, CourseGroup, GroupAccess, Material


class PlatformTests(TestCase):
    @classmethod
    def setUpTestData(cls):
        call_command("seed_demo", stdout=open("/dev/null", "w"))
        cls.admin = User.objects.get(username="admin")
        cls.emp = User.objects.get(username="employee")
        cls.cand = User.objects.get(username="candidate")

    def test_login_required(self):
        r = self.client.get(reverse("home"))
        self.assertRedirects(r, reverse("login") + "?next=/")

    def test_employee_cannot_open_admin(self):
        self.client.force_login(self.emp)
        self.assertEqual(self.client.get(reverse("manage:dashboard")).status_code, 403)
        self.assertEqual(self.client.get(reverse("manage:user_list")).status_code, 403)

    def test_employee_sees_only_accessible_courses(self):
        self.client.force_login(self.emp)
        r = self.client.get(reverse("courses"))
        self.assertContains(r, "Знакомство с компанией")
        self.assertContains(r, "Дезинфекция: основы")
        self.client.force_login(self.cand)
        r = self.client.get(reverse("courses"))
        self.assertNotContains(r, "Знакомство с компанией")
        course = Course.objects.get(title="Знакомство с компанией")
        self.assertEqual(self.client.get(reverse("course_detail", args=[course.pk])).status_code, 404)

    def test_expired_and_unpublished_access(self):
        group = CourseGroup.objects.get(title="Адаптация")
        GroupAccess.objects.filter(user=self.emp, group=group).update(expires_at=timezone.now() - timedelta(days=1))
        self.client.force_login(self.emp)
        self.assertNotContains(self.client.get(reverse("courses")), "Знакомство с компанией")
        GroupAccess.objects.filter(user=self.emp, group=group).update(expires_at=None)
        Course.objects.filter(title="Знакомство с компанией").update(is_published=False)
        self.assertNotContains(self.client.get(reverse("courses")), "Знакомство с компанией")

    def test_material_progress(self):
        self.client.force_login(self.emp)
        m = Material.objects.filter(course__title="Знакомство с компанией").first()
        self.client.post(reverse("material_complete", args=[m.pk]))
        self.assertEqual(m.course.progress_for(self.emp)["done"], 1)

    def _take(self, user, test, answers_fn):
        a = Assignment.objects.get(user=user, test=test)
        self.client.force_login(user)
        r = self.client.post(reverse("test_start", args=[a.pk]))
        attempt = Attempt.objects.get(assignment=a)
        self.assertRedirects(r, reverse("attempt_take", args=[attempt.pk]))
        self.assertEqual(self.client.get(r.url).status_code, 200)
        data = answers_fn(attempt)
        self.client.post(reverse("attempt_submit", args=[attempt.pk]), data)
        attempt.refresh_from_db()
        return a, attempt

    def test_candidate_results_hidden_by_default(self):
        test = Test.objects.get(title="Входной тест для кандидатов")

        def correct(attempt):
            q = test.questions.first()
            return {f"q{q.id}": [str(q.choices.get(is_correct=True).id)]}

        a, attempt = self._take(self.cand, test, correct)
        self.assertTrue(attempt.passed)
        self.assertEqual(attempt.percent, 100)
        r = self.client.get(reverse("attempt_done", args=[attempt.pk]))
        self.assertContains(r, "Ответы отправлены")
        self.assertNotContains(r, "100%")
        # Одноразовый доступ: повторно начать нельзя.
        self.client.post(reverse("test_start", args=[a.pk]))
        self.assertEqual(Attempt.objects.filter(assignment=a).count(), 1)
        # Админ видит результат.
        self.client.force_login(self.admin)
        self.assertContains(self.client.get(reverse("manage:attempt_detail", args=[attempt.pk])), "100%")

    def test_results_shown_when_enabled_and_manual_grading(self):
        test = Test.objects.get(title="Тест: охрана труда")
        test.show_results = True
        test.save()

        def answers(attempt):
            data = {}
            for q in test.questions.all():
                if q.kind == Question.Kind.TEXT:
                    data[f"q{q.id}"] = "Промыть водой"
                else:
                    data[f"q{q.id}"] = [str(c.id) for c in q.choices.filter(is_correct=True)]
            return data

        _, attempt = self._take(self.emp, test, answers)
        self.assertTrue(attempt.needs_review)
        self.assertIsNone(attempt.passed)
        self.assertContains(self.client.get(reverse("attempt_done", args=[attempt.pk])), "проверяет администратор")
        self.client.force_login(self.admin)
        ans = attempt.answers.get(question__kind=Question.Kind.TEXT)
        self.client.post(reverse("manage:attempt_detail", args=[attempt.pk]), {f"a{ans.id}": "2"})
        attempt.refresh_from_db()
        self.assertFalse(attempt.needs_review)
        self.assertTrue(attempt.passed)
        self.assertEqual(attempt.percent, 100)
        self.client.force_login(self.emp)
        self.assertContains(self.client.get(reverse("attempt_done", args=[attempt.pk])), "Тест сдан")

    def test_timeout_finishes_attempt(self):
        test = Test.objects.get(title="Входной тест для кандидатов")
        a = Assignment.objects.get(user=self.cand, test=test)
        attempt = a.start_attempt()
        Attempt.objects.filter(pk=attempt.pk).update(deadline=timezone.now() - timedelta(minutes=5))
        self.client.force_login(self.cand)
        r = self.client.get(reverse("attempt_take", args=[attempt.pk]))
        self.assertRedirects(r, reverse("attempt_done", args=[attempt.pk]))
        attempt.refresh_from_db()
        self.assertEqual(attempt.status, Attempt.Status.TIMED_OUT)

    def test_cannot_open_foreign_attempt(self):
        test = Test.objects.get(title="Входной тест для кандидатов")
        attempt = Assignment.objects.get(user=self.cand, test=test).start_attempt()
        self.client.force_login(self.emp)
        self.assertEqual(self.client.get(reverse("attempt_take", args=[attempt.pk])).status_code, 404)

    def test_require_course_completion(self):
        test = Test.objects.get(title="Тест: охрана труда")
        test.require_course_completion = True
        test.save()
        a = Assignment.objects.get(user=self.emp, test=test)
        self.assertEqual(a.state(), "locked")
        self.client.force_login(self.emp)
        for m in test.course.materials.all():
            self.client.post(reverse("material_complete", args=[m.pk]))
        self.assertEqual(a.state(), "available")

    def test_admin_can_create_user_and_grant_access(self):
        self.client.force_login(self.admin)
        r = self.client.post(reverse("manage:user_create"), {
            "last_name": "Иванов", "first_name": "Иван", "username": "ivanov", "role": "employee",
            "is_active": "on", "password1": "Strong-Pass-123", "password2": "Strong-Pass-123",
        })
        u = User.objects.get(username="ivanov")
        self.assertRedirects(r, reverse("manage:user_access", args=[u.pk]))
        group = CourseGroup.objects.get(title="Адаптация")
        test = Test.objects.get(title="Входной тест для кандидатов")
        self.client.post(reverse("manage:user_access", args=[u.pk]), {"groups": [group.pk], "tests": [test.pk]})
        self.assertTrue(GroupAccess.objects.filter(user=u, group=group).exists())
        self.assertTrue(Assignment.objects.filter(user=u, test=test).exists())

    def test_question_validation(self):
        self.client.force_login(self.admin)
        test = Test.objects.get(title="Входной тест для кандидатов")
        base = {
            "kind": "single", "text": "Новый?", "points": 1, "order": 5,
            "choices-TOTAL_FORMS": 2, "choices-INITIAL_FORMS": 0, "choices-MIN_NUM_FORMS": 0, "choices-MAX_NUM_FORMS": 1000,
            "choices-0-text": "A", "choices-0-is_correct": "on", "choices-1-text": "B",
        }
        self.client.post(reverse("manage:question_create", args=[test.pk]), base)
        q = test.questions.get(text="Новый?")
        self.assertEqual(q.choices.count(), 2)
        bad = dict(base, text="Плохой", **{"choices-1-is_correct": "on"})
        self.client.post(reverse("manage:question_create", args=[test.pk]), bad)
        self.assertFalse(test.questions.filter(text="Плохой").exists())

    def test_all_admin_pages_render(self):
        self.client.force_login(self.admin)
        g = CourseGroup.objects.first()
        c = Course.objects.first()
        m = Material.objects.first()
        t = Test.objects.first()
        q = Question.objects.first()
        a = Assignment.objects.first()
        att = a.start_attempt()
        att.finish()
        urls = [
            reverse("manage:dashboard"), reverse("manage:group_list"), reverse("manage:group_create"),
            reverse("manage:group_detail", args=[g.pk]), reverse("manage:group_edit", args=[g.pk]),
            reverse("manage:group_access", args=[g.pk]), reverse("manage:group_delete", args=[g.pk]),
            reverse("manage:course_list"), reverse("manage:course_create"),
            reverse("manage:course_detail", args=[c.pk]), reverse("manage:course_edit", args=[c.pk]),
            reverse("manage:course_access", args=[c.pk]), reverse("manage:material_create", args=[c.pk]),
            reverse("manage:material_edit", args=[m.pk]), reverse("manage:test_list"), reverse("manage:test_create"),
            reverse("manage:test_detail", args=[t.pk]), reverse("manage:test_edit", args=[t.pk]),
            reverse("manage:assign", args=[t.pk]), reverse("manage:question_create", args=[t.pk]),
            reverse("manage:question_edit", args=[q.pk]), reverse("manage:assign_edit", args=[a.pk]),
            reverse("manage:results"), reverse("manage:results_export"), reverse("manage:attempt_detail", args=[att.pk]),
            reverse("manage:user_list"), reverse("manage:user_create"),
            reverse("manage:user_detail", args=[self.emp.pk]), reverse("manage:user_edit", args=[self.emp.pk]),
            reverse("manage:user_password", args=[self.emp.pk]), reverse("manage:user_access", args=[self.emp.pk]),
            reverse("material_detail", args=[m.pk]), reverse("course_detail", args=[c.pk]), reverse("password_change"),
        ]
        for url in urls:
            with self.subTest(url=url):
                self.assertEqual(self.client.get(url).status_code, 200)

    def test_employee_pages_render(self):
        self.client.force_login(self.emp)
        a = Assignment.objects.filter(user=self.emp).first()
        for url in [reverse("home"), reverse("courses"), reverse("my_tests"), reverse("test_intro", args=[a.pk])]:
            with self.subTest(url=url):
                self.assertEqual(self.client.get(url).status_code, 200)


class ImportTests(TestCase):
    def test_import_example_content(self):
        import tempfile
        from pathlib import Path

        from django.test import override_settings

        from exams.models import Test
        from learning.importer import Importer, parse_test

        root = Path(__file__).resolve().parent.parent / "content_example"
        with tempfile.TemporaryDirectory() as media, override_settings(MEDIA_ROOT=media):
            report = Importer(root).run()
            self.assertEqual(report.errors, [])
            group = CourseGroup.objects.get(title="Адаптация")
            self.assertEqual(group.courses.count(), 2)
            self.assertTrue(Material.objects.filter(title="СИЗ", body__contains="<table>").exists())
            self.assertEqual(Material.objects.get(title="Приветствие директора").kind, Material.Kind.VIDEO)
            self.assertTrue(Material.objects.get(title="Инструкция").file.name.endswith(".pdf"))
            # «Тестирование оборудования.txt» — материал, а не тест
            self.assertTrue(Material.objects.filter(title="Тестирование оборудования").exists())
            t = Test.objects.get(title="Тест: знакомство с компанией")
            self.assertEqual(t.time_limit_minutes, 10)
            self.assertEqual([q.kind for q in t.questions.all()], ["single", "multiple", "text"])
            self.assertEqual(Test.objects.get(title="Входной тест для кандидатов").course, None)
            # Повторный импорт ничего не дублирует и не трогает правки на сайте.
            Material.objects.filter(title="О компании").update(body="Отредактировано на сайте")
            again = Importer(root).run()
            self.assertEqual(again.created, [])
            self.assertEqual(Material.objects.get(title="О компании").body, "Отредактировано на сайте")
            self.assertEqual(Material.objects.count(), 5)

        spec = parse_test("? Вопрос\n* да\n- нет\nБаллы: 2")
        self.assertEqual(spec["questions"][0]["points"], 2)
        self.assertEqual(spec["questions"][0]["choices"], [("да", True), ("нет", False)])
