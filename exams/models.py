import random
from decimal import Decimal

from django.conf import settings
from django.db import models, transaction
from django.utils import timezone

from learning.models import Course

GRACE_SECONDS = 30


class Test(models.Model):
    title = models.CharField("Название", max_length=200)
    description = models.TextField("Описание / инструкция", blank=True)
    course = models.ForeignKey(
        Course, verbose_name="Курс", related_name="tests", on_delete=models.SET_NULL, null=True, blank=True,
        help_text="Необязательно. Оставьте пустым для отдельного теста (например, для кандидатов).",
    )
    time_limit_minutes = models.PositiveIntegerField(
        "Ограничение по времени, мин", null=True, blank=True, help_text="Пусто — без ограничения."
    )
    pass_percent = models.PositiveSmallIntegerField("Проходной балл, %", default=70)
    questions_per_attempt = models.PositiveIntegerField(
        "Вопросов в попытке", null=True, blank=True,
        help_text="Пусто — все вопросы. Иначе — случайная выборка указанного размера.",
    )
    shuffle_questions = models.BooleanField("Перемешивать вопросы", default=True)
    shuffle_choices = models.BooleanField("Перемешивать варианты ответов", default=True)
    show_results = models.BooleanField(
        "Показывать результат пользователю", default=False,
        help_text="По умолчанию результат видит только администратор.",
    )
    show_correct_answers = models.BooleanField(
        "Показывать правильные ответы", default=False,
        help_text="Работает только вместе с «Показывать результат».",
    )
    require_course_completion = models.BooleanField(
        "Только после изучения курса", default=False,
        help_text="Тест откроется, когда пользователь изучит все материалы привязанного курса.",
    )
    is_active = models.BooleanField("Активен", default=True)
    source_hash = models.CharField(max_length=64, blank=True, editable=False)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, related_name="+", on_delete=models.SET_NULL, null=True, blank=True
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Тест"
        verbose_name_plural = "Тесты"
        ordering = ["-created_at"]

    def __str__(self):
        return self.title

    @property
    def question_count(self):
        n = self.questions.count()
        if self.questions_per_attempt:
            return min(n, self.questions_per_attempt)
        return n


class Question(models.Model):
    class Kind(models.TextChoices):
        SINGLE = "single", "Один правильный ответ"
        MULTIPLE = "multiple", "Несколько правильных ответов"
        TEXT = "text", "Развёрнутый ответ (проверяет администратор)"

    test = models.ForeignKey(Test, related_name="questions", on_delete=models.CASCADE)
    kind = models.CharField("Тип вопроса", max_length=16, choices=Kind.choices, default=Kind.SINGLE)
    text = models.TextField("Вопрос")
    image = models.FileField("Изображение", upload_to="questions/%Y/%m/", blank=True)
    points = models.PositiveSmallIntegerField("Баллы", default=1)
    explanation = models.TextField("Пояснение к ответу", blank=True)
    order = models.PositiveIntegerField("Порядок", default=0)

    class Meta:
        verbose_name = "Вопрос"
        verbose_name_plural = "Вопросы"
        ordering = ["order", "id"]

    def __str__(self):
        return self.text[:80]

    @property
    def is_text(self):
        return self.kind == self.Kind.TEXT


class Choice(models.Model):
    question = models.ForeignKey(Question, related_name="choices", on_delete=models.CASCADE)
    text = models.CharField("Вариант ответа", max_length=500)
    is_correct = models.BooleanField("Верный", default=False)
    order = models.PositiveIntegerField("Порядок", default=0)

    class Meta:
        ordering = ["order", "id"]

    def __str__(self):
        return self.text


class Assignment(models.Model):
    """Разовый доступ пользователя к сдаче теста."""

    test = models.ForeignKey(Test, related_name="assignments", on_delete=models.CASCADE)
    user = models.ForeignKey(settings.AUTH_USER_MODEL, related_name="test_assignments", on_delete=models.CASCADE)
    attempts_allowed = models.PositiveSmallIntegerField("Количество попыток", default=1)
    available_from = models.DateTimeField("Доступен с", null=True, blank=True)
    available_until = models.DateTimeField("Доступен до", null=True, blank=True)
    is_revoked = models.BooleanField("Отозван", default=False)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, related_name="+", on_delete=models.SET_NULL, null=True, blank=True
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Доступ к тесту"
        verbose_name_plural = "Доступы к тестам"
        ordering = ["-created_at"]

    def __str__(self):
        return f"{self.test} → {self.user}"

    @property
    def attempts_used(self):
        return self.attempts.count()

    @property
    def attempts_left(self):
        return max(self.attempts_allowed - self.attempts_used, 0)

    def current_attempt(self):
        return self.attempts.filter(status=Attempt.Status.IN_PROGRESS).first()

    def last_finished_attempt(self):
        return self.attempts.exclude(status=Attempt.Status.IN_PROGRESS).order_by("-finished_at").first()

    def course_ready(self):
        t = self.test
        if not (t.require_course_completion and t.course_id):
            return True
        return t.course.is_completed_by(self.user)

    def state(self):
        """Состояние для отображения: available / in_progress / scheduled / expired / done / revoked / locked."""
        now = timezone.now()
        if self.is_revoked or not self.test.is_active:
            return "revoked"
        if self.current_attempt():
            return "in_progress"
        if self.attempts_left == 0:
            return "done"
        if self.available_from and now < self.available_from:
            return "scheduled"
        if self.available_until and now > self.available_until:
            return "done" if self.attempts_used else "expired"
        if not self.course_ready():
            return "locked"
        return "available"

    def can_start(self):
        return self.state() == "available"

    @transaction.atomic
    def start_attempt(self):
        questions = list(self.test.questions.prefetch_related("choices"))
        if self.test.shuffle_questions:
            random.shuffle(questions)
        if self.test.questions_per_attempt:
            if not self.test.shuffle_questions:
                questions = random.sample(questions, min(len(questions), self.test.questions_per_attempt))
                questions.sort(key=lambda q: (q.order, q.id))
            else:
                questions = questions[: self.test.questions_per_attempt]
        choice_order = {}
        for q in questions:
            ids = [c.id for c in q.choices.all()]
            if self.test.shuffle_choices:
                random.shuffle(ids)
            choice_order[str(q.id)] = ids
        now = timezone.now()
        deadline = None
        if self.test.time_limit_minutes:
            deadline = now + timezone.timedelta(minutes=self.test.time_limit_minutes)
        if self.available_until and (deadline is None or deadline > self.available_until):
            deadline = self.available_until
        return Attempt.objects.create(
            assignment=self, user=self.user, test=self.test, started_at=now, deadline=deadline,
            question_order=[q.id for q in questions], choice_order=choice_order,
            max_score=sum(q.points for q in questions),
        )


class Attempt(models.Model):
    class Status(models.TextChoices):
        IN_PROGRESS = "in_progress", "В процессе"
        SUBMITTED = "submitted", "Завершён"
        TIMED_OUT = "timed_out", "Время вышло"

    assignment = models.ForeignKey(Assignment, related_name="attempts", on_delete=models.CASCADE)
    user = models.ForeignKey(settings.AUTH_USER_MODEL, related_name="attempts", on_delete=models.CASCADE)
    test = models.ForeignKey(Test, related_name="attempts", on_delete=models.CASCADE)
    status = models.CharField(max_length=16, choices=Status.choices, default=Status.IN_PROGRESS)
    started_at = models.DateTimeField(default=timezone.now)
    deadline = models.DateTimeField(null=True, blank=True)
    finished_at = models.DateTimeField(null=True, blank=True)
    question_order = models.JSONField(default=list)
    choice_order = models.JSONField(default=dict)
    score = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    max_score = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    percent = models.DecimalField(max_digits=5, decimal_places=1, default=0)
    passed = models.BooleanField(null=True)
    needs_review = models.BooleanField(default=False)

    class Meta:
        verbose_name = "Попытка"
        verbose_name_plural = "Попытки"
        ordering = ["-started_at"]

    def __str__(self):
        return f"{self.user} — {self.test}"

    @property
    def is_finished(self):
        return self.status != self.Status.IN_PROGRESS

    def is_overdue(self, grace=True):
        if not self.deadline:
            return False
        extra = timezone.timedelta(seconds=GRACE_SECONDS if grace else 0)
        return timezone.now() > self.deadline + extra

    @property
    def seconds_left(self):
        if not self.deadline:
            return None
        return max(int((self.deadline - timezone.now()).total_seconds()), 0)

    @property
    def duration(self):
        if not self.finished_at:
            return None
        return self.finished_at - self.started_at

    def ordered_questions(self):
        qs = {q.id: q for q in Question.objects.filter(id__in=self.question_order).prefetch_related("choices")}
        result = []
        for qid in self.question_order:
            q = qs.get(qid)
            if not q:
                continue
            choices = {c.id: c for c in q.choices.all()}
            order = self.choice_order.get(str(qid), list(choices))
            q.ordered_choices = [choices[c] for c in order if c in choices]
            q.ordered_choices += [c for cid, c in choices.items() if cid not in order]
            result.append(q)
        return result

    def save_answers(self, data):
        """data — QueryDict из формы прохождения теста."""
        for q in self.ordered_questions():
            answer, _ = Answer.objects.get_or_create(attempt=self, question=q)
            if q.is_text:
                answer.text_answer = data.get(f"q{q.id}", "")[:10000]
                answer.save(update_fields=["text_answer"])
            else:
                valid = {c.id for c in q.ordered_choices}
                raw = data.getlist(f"q{q.id}")
                ids = []
                for v in raw:
                    try:
                        if int(v) in valid:
                            ids.append(int(v))
                    except (TypeError, ValueError):
                        pass
                if q.kind == Question.Kind.SINGLE:
                    ids = ids[:1]
                answer.selected.set(ids)

    @transaction.atomic
    def finish(self, timed_out=False):
        if self.is_finished:
            return
        self.status = self.Status.TIMED_OUT if timed_out else self.Status.SUBMITTED
        self.finished_at = timezone.now()
        self.grade()

    def grade(self):
        answers = {a.question_id: a for a in self.answers.prefetch_related("selected")}
        score = Decimal(0)
        needs_review = False
        for q in self.ordered_questions():
            a = answers.get(q.id) or Answer.objects.create(attempt=self, question=q)
            if q.is_text:
                if a.points is None:
                    if a.text_answer.strip():
                        needs_review = True
                    else:
                        a.points = Decimal(0)
                        a.is_correct = False
                        a.save(update_fields=["points", "is_correct"])
                score += a.points or 0
                continue
            correct = {c.id for c in q.ordered_choices if c.is_correct}
            chosen = {c.id for c in a.selected.all()}
            ok = bool(correct) and chosen == correct
            a.is_correct = ok
            a.points = Decimal(q.points) if ok else Decimal(0)
            a.save(update_fields=["is_correct", "points"])
            score += a.points
        self.score = score
        self.needs_review = needs_review
        self.percent = round(score * 100 / self.max_score, 1) if self.max_score else Decimal(0)
        self.passed = None if needs_review else self.percent >= self.test.pass_percent
        self.save()

    @property
    def result_visible_to_user(self):
        return self.test.show_results and self.is_finished and not self.needs_review


class Answer(models.Model):
    attempt = models.ForeignKey(Attempt, related_name="answers", on_delete=models.CASCADE)
    question = models.ForeignKey(Question, related_name="answers", on_delete=models.CASCADE)
    selected = models.ManyToManyField(Choice, blank=True, related_name="+")
    text_answer = models.TextField(blank=True)
    is_correct = models.BooleanField(null=True)
    points = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)

    class Meta:
        constraints = [models.UniqueConstraint(fields=["attempt", "question"], name="uniq_attempt_answer")]
