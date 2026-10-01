import csv

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.db.models import Avg, Count, Q
from django.http import HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.http import require_POST

from accounts.decorators import admin_required
from accounts.models import User

from .forms import (
    AssignForm, AssignmentEditForm, ChoiceFormSet, GradeForm, NewChoiceFormSet, QuestionForm, TestForm,
)
from .models import Answer, Assignment, Attempt, Question, Test

STATE_LABELS = {
    "available": ("Доступен", "blue"),
    "in_progress": ("Начат", "orange"),
    "scheduled": ("Запланирован", "gray"),
    "locked": ("Сначала изучите курс", "gray"),
    "expired": ("Срок истёк", "red"),
    "done": ("Сдан", "green"),
    "revoked": ("Отозван", "gray"),
}


def _decorate(assignment):
    state = assignment.state()
    assignment.state_code = state
    assignment.state_label, assignment.state_color = STATE_LABELS[state]
    return assignment


# ---------- Сотрудник ----------


@login_required
def my_tests(request):
    assignments = [
        _decorate(a) for a in Assignment.objects.filter(user=request.user, test__is_active=True)
        .exclude(is_revoked=True).select_related("test", "test__course").prefetch_related("attempts")
    ]
    active = [a for a in assignments if a.state_code in ("available", "in_progress", "scheduled", "locked")]
    finished = []
    for a in assignments:
        for att in a.attempts.all():
            if att.is_finished:
                finished.append(att)
    finished.sort(key=lambda x: x.finished_at, reverse=True)
    expired = [a for a in assignments if a.state_code == "expired"]
    return render(request, "employee/tests.html", {"active": active, "finished": finished, "expired": expired})


def _own_assignment(request, pk):
    return get_object_or_404(Assignment.objects.select_related("test", "test__course"), pk=pk, user=request.user)


@login_required
def test_intro(request, pk):
    a = _decorate(_own_assignment(request, pk))
    return render(request, "employee/test_intro.html", {"a": a, "test": a.test})


@login_required
@require_POST
def test_start(request, pk):
    with transaction.atomic():
        a = Assignment.objects.select_for_update().get(pk=_own_assignment(request, pk).pk)
        current = a.current_attempt()
        if current:
            return redirect("attempt_take", current.pk)
        if not a.can_start():
            messages.error(request, "Этот тест сейчас недоступен.")
            return redirect("test_intro", a.pk)
        if not a.test.questions.exists():
            messages.error(request, "В тесте пока нет вопросов. Обратитесь к администратору.")
            return redirect("test_intro", a.pk)
        attempt = a.start_attempt()
    return redirect("attempt_take", attempt.pk)


def _own_attempt(request, pk):
    return get_object_or_404(Attempt.objects.select_related("test", "assignment"), pk=pk, user=request.user)


@login_required
def attempt_take(request, pk):
    attempt = _own_attempt(request, pk)
    if attempt.is_finished:
        return redirect("attempt_done", attempt.pk)
    if attempt.is_overdue():
        attempt.finish(timed_out=True)
        return redirect("attempt_done", attempt.pk)
    questions = attempt.ordered_questions()
    answers = {a.question_id: a for a in attempt.answers.prefetch_related("selected")}
    for q in questions:
        a = answers.get(q.id)
        q.selected_ids = {c.id for c in a.selected.all()} if a else set()
        q.text_value = a.text_answer if a else ""
    return render(request, "employee/attempt_take.html", {
        "attempt": attempt, "test": attempt.test, "questions": questions,
    })


@login_required
@require_POST
def attempt_save(request, pk):
    attempt = _own_attempt(request, pk)
    done_url = reverse("attempt_done", args=[attempt.pk])
    if attempt.is_finished:
        return JsonResponse({"finished": True, "redirect": done_url})
    if attempt.is_overdue():
        attempt.finish(timed_out=True)
        return JsonResponse({"finished": True, "redirect": done_url})
    attempt.save_answers(request.POST)
    return JsonResponse({"ok": True, "seconds_left": attempt.seconds_left})


@login_required
@require_POST
def attempt_submit(request, pk):
    attempt = _own_attempt(request, pk)
    if not attempt.is_finished:
        overdue = attempt.is_overdue()
        if not overdue:
            attempt.save_answers(request.POST)
        attempt.finish(timed_out=overdue or attempt.is_overdue(grace=False))
    return redirect("attempt_done", attempt.pk)


@login_required
def attempt_done(request, pk):
    attempt = _own_attempt(request, pk)
    if not attempt.is_finished:
        return redirect("attempt_take", attempt.pk)
    questions = []
    if attempt.result_visible_to_user and attempt.test.show_correct_answers:
        questions = _review_questions(attempt)
    return render(request, "employee/attempt_done.html", {"attempt": attempt, "questions": questions})


def _review_questions(attempt):
    questions = attempt.ordered_questions()
    answers = {a.question_id: a for a in attempt.answers.prefetch_related("selected")}
    for q in questions:
        a = answers.get(q.id)
        q.answer = a
        q.selected_ids = {c.id for c in a.selected.all()} if a else set()
    return questions


# ---------- Администратор: тесты ----------


@admin_required
def test_list(request):
    tests = Test.objects.select_related("course").annotate(
        n_questions=Count("questions", distinct=True),
        n_assigned=Count("assignments", distinct=True),
        n_attempts=Count("attempts", filter=~Q(attempts__status=Attempt.Status.IN_PROGRESS), distinct=True),
    )
    return render(request, "manage/test_list.html", {"tests": tests})


@admin_required
def test_detail(request, pk):
    test = get_object_or_404(Test.objects.select_related("course"), pk=pk)
    finished = test.attempts.exclude(status=Attempt.Status.IN_PROGRESS)
    stats = finished.aggregate(avg=Avg("percent"), n=Count("id"), passed=Count("id", filter=Q(passed=True)))
    assignments = [_decorate(a) for a in test.assignments.select_related("user").prefetch_related("attempts")]
    return render(request, "manage/test_detail.html", {
        "test": test,
        "questions": test.questions.prefetch_related("choices"),
        "assignments": assignments,
        "stats": stats,
        "total_points": sum(q.points for q in test.questions.all()),
    })


@admin_required
def test_form(request, pk=None):
    test = get_object_or_404(Test, pk=pk) if pk else None
    initial = {"course": request.GET.get("course")} if not pk and request.GET.get("course") else {}
    form = TestForm(request.POST or None, instance=test, initial=initial)
    if request.method == "POST" and form.is_valid():
        t = form.save(commit=False)
        if not t.pk:
            t.created_by = request.user
        t.save()
        messages.success(request, "Тест сохранён." if pk else "Тест создан. Добавьте вопросы.")
        return redirect("manage:test_detail", t.pk)
    return render(request, "manage/form.html", {
        "form": form, "title": "Настройки теста" if pk else "Новый тест",
        "back_url": "manage:test_detail" if pk else "manage:test_list", "back_arg": pk,
    })


@admin_required
def test_delete(request, pk):
    test = get_object_or_404(Test, pk=pk)
    if request.method == "POST":
        test.delete()
        messages.success(request, "Тест удалён.")
        return redirect("manage:test_list")
    return render(request, "manage/confirm_delete.html", {
        "object": test, "kind": "тест", "warning": "Будут удалены все вопросы, доступы и результаты по тесту.",
        "back_url": "manage:test_detail", "back_arg": pk,
    })


@admin_required
def question_form(request, test_pk=None, pk=None):
    question = get_object_or_404(Question, pk=pk) if pk else None
    test = question.test if question else get_object_or_404(Test, pk=test_pk)
    initial = {}
    if not question:
        last = test.questions.order_by("-order").first()
        initial["order"] = (last.order + 1) if last else 1
    form = QuestionForm(request.POST or None, request.FILES or None, instance=question, initial=initial)
    FormSet = ChoiceFormSet if question else NewChoiceFormSet
    formset = FormSet(request.POST or None, instance=question or Question(test=test), prefix="choices")
    if request.method == "POST" and form.is_valid() and formset.is_valid():
        kind = form.cleaned_data["kind"]
        kept = [f for f in formset.forms if f.cleaned_data and not f.cleaned_data.get("DELETE")
                and f.cleaned_data.get("text")]
        n_correct = sum(1 for f in kept if f.cleaned_data.get("is_correct"))
        error = None
        if kind != Question.Kind.TEXT:
            if len(kept) < 2:
                error = "Добавьте минимум два варианта ответа."
            elif n_correct == 0:
                error = "Отметьте хотя бы один верный вариант."
            elif kind == Question.Kind.SINGLE and n_correct > 1:
                error = "Для вопроса с одним ответом верным может быть только один вариант."
        if error:
            messages.error(request, error)
        else:
            with transaction.atomic():
                q = form.save(commit=False)
                q.test = test
                q.save()
                formset.instance = q
                formset.save(commit=False)
                for obj in formset.deleted_objects:
                    obj.delete()
                if kind == Question.Kind.TEXT:
                    q.choices.all().delete()
                else:
                    for i, f in enumerate(kept, start=1):
                        f.instance.question = q
                        f.instance.order = i
                        f.instance.save()
            messages.success(request, "Вопрос сохранён.")
            if request.POST.get("add_another"):
                return redirect("manage:question_create", test.pk)
            return redirect("manage:test_detail", test.pk)
    return render(request, "manage/question_form.html", {
        "form": form, "formset": formset, "test": test, "question": question,
    })


@admin_required
def question_delete(request, pk):
    q = get_object_or_404(Question, pk=pk)
    if request.method == "POST":
        test_pk = q.test_id
        q.delete()
        messages.success(request, "Вопрос удалён.")
        return redirect("manage:test_detail", test_pk)
    return render(request, "manage/confirm_delete.html", {
        "object": q, "kind": "вопрос", "back_url": "manage:test_detail", "back_arg": q.test_id,
    })


# ---------- Администратор: разовые доступы ----------


@admin_required
def assign(request, pk):
    test = get_object_or_404(Test, pk=pk)
    form = AssignForm(request.POST or None)
    users = User.objects.filter(is_active=True).exclude(role=User.Role.ADMIN).exclude(is_superuser=True)
    open_ids = {
        a.user_id for a in test.assignments.filter(is_revoked=False).prefetch_related("attempts")
        if a.state() in ("available", "in_progress", "scheduled", "locked")
    }
    if request.method == "POST" and form.is_valid():
        selected = {int(x) for x in request.POST.getlist("users")}
        if not selected:
            messages.error(request, "Выберите хотя бы одного пользователя.")
        else:
            created = 0
            for uid in selected - open_ids:
                Assignment.objects.create(test=test, user_id=uid, created_by=request.user, **form.cleaned_data)
                created += 1
            skipped = len(selected & open_ids)
            msg = f"Доступ выдан: {created}."
            if skipped:
                msg += f" Пропущено (уже есть открытый доступ): {skipped}."
            messages.success(request, msg)
            return redirect("manage:test_detail", test.pk)
    return render(request, "manage/assign.html", {
        "test": test, "form": form, "users": users, "open_ids": open_ids,
        "preselect": {int(x) for x in request.GET.getlist("user") if x.isdigit()},
    })


@admin_required
def assignment_edit(request, pk):
    a = get_object_or_404(Assignment.objects.select_related("test", "user"), pk=pk)
    form = AssignmentEditForm(request.POST or None, instance=a)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Доступ обновлён.")
        return redirect("manage:test_detail", a.test_id)
    return render(request, "manage/form.html", {
        "form": form, "title": f"Доступ: {a.user} — {a.test}", "back_url": "manage:test_detail", "back_arg": a.test_id,
    })


@admin_required
@require_POST
def assignment_extra_attempt(request, pk):
    a = get_object_or_404(Assignment, pk=pk)
    a.attempts_allowed += 1
    a.is_revoked = False
    a.save(update_fields=["attempts_allowed", "is_revoked"])
    messages.success(request, f"Пользователю «{a.user}» добавлена попытка.")
    return redirect(request.POST.get("back") or reverse("manage:test_detail", args=[a.test_id]))


@admin_required
@require_POST
def assignment_revoke(request, pk):
    a = get_object_or_404(Assignment, pk=pk)
    if a.attempts.exists():
        a.is_revoked = True
        a.save(update_fields=["is_revoked"])
        cur = a.current_attempt()
        if cur:
            cur.finish()
        messages.success(request, "Доступ отозван. Результаты сохранены.")
    else:
        a.delete()
        messages.success(request, "Доступ удалён.")
    return redirect(request.POST.get("back") or reverse("manage:test_detail", args=[a.test_id]))


# ---------- Администратор: результаты ----------


def _filtered_attempts(request):
    attempts = Attempt.objects.exclude(status=Attempt.Status.IN_PROGRESS).select_related("user", "test")
    f = {k: request.GET.get(k, "") for k in ("test", "kind", "status", "q")}
    if f["test"].isdigit():
        attempts = attempts.filter(test_id=int(f["test"]))
    if f["kind"] == "candidate":
        attempts = attempts.filter(user__is_candidate=True)
    elif f["kind"] == "employee":
        attempts = attempts.filter(user__is_candidate=False)
    if f["status"] == "review":
        attempts = attempts.filter(needs_review=True)
    elif f["status"] == "passed":
        attempts = attempts.filter(passed=True)
    elif f["status"] == "failed":
        attempts = attempts.filter(passed=False)
    if f["q"]:
        q = f["q"]
        attempts = attempts.filter(
            Q(user__last_name__icontains=q) | Q(user__first_name__icontains=q) | Q(user__username__icontains=q)
        )
    return attempts, f


@admin_required
def results(request):
    attempts, f = _filtered_attempts(request)
    return render(request, "manage/results.html", {
        "attempts": attempts[:500], "f": f, "tests": Test.objects.all(),
        "in_progress": Attempt.objects.filter(status=Attempt.Status.IN_PROGRESS).select_related("user", "test"),
        "query": request.GET.urlencode(),
    })


@admin_required
def results_export(request):
    attempts, _ = _filtered_attempts(request)
    resp = HttpResponse(content_type="text/csv; charset=utf-8")
    resp["Content-Disposition"] = f'attachment; filename="results-{timezone.localdate():%Y-%m-%d}.csv"'
    resp.write("﻿")
    w = csv.writer(resp, delimiter=";")
    w.writerow(["ФИО", "Логин", "Тип", "Тест", "Начало", "Окончание", "Баллы", "Макс.", "%", "Итог"])
    for a in attempts:
        verdict = "На проверке" if a.needs_review else ("Сдан" if a.passed else "Не сдан")
        w.writerow([
            a.user.get_full_name(), a.user.username, a.user.kind_label, a.test.title,
            timezone.localtime(a.started_at).strftime("%d.%m.%Y %H:%M"),
            timezone.localtime(a.finished_at).strftime("%d.%m.%Y %H:%M") if a.finished_at else "",
            a.score, a.max_score, a.percent, verdict,
        ])
    return resp


@admin_required
def attempt_detail(request, pk):
    attempt = get_object_or_404(Attempt.objects.select_related("user", "test", "assignment"), pk=pk)
    questions = _review_questions(attempt)
    text_answers = [q.answer for q in questions if q.is_text and q.answer]
    for a in text_answers:
        a.question = next(q for q in questions if q.id == a.question_id)
    form = GradeForm(request.POST or None, answers=text_answers)
    if request.method == "POST" and attempt.is_finished and form.is_valid():
        for a in text_answers:
            pts = form.cleaned_data[a.field_name]
            Answer.objects.filter(pk=a.pk).update(points=pts, is_correct=pts >= a.question.points)
        attempt.grade()
        messages.success(request, "Оценки сохранены, результат пересчитан.")
        return redirect("manage:attempt_detail", attempt.pk)
    for q in questions:
        q.grade_field = form[q.answer.field_name] if q.is_text and q.answer and hasattr(q.answer, "field_name") else None
    return render(request, "manage/attempt_detail.html", {"attempt": attempt, "questions": questions, "form": form})


@admin_required
def attempt_delete(request, pk):
    attempt = get_object_or_404(Attempt.objects.select_related("user", "test"), pk=pk)
    if request.method == "POST":
        attempt.delete()
        messages.success(request, "Попытка удалена — у пользователя освободилась попытка.")
        return redirect("manage:results")
    return render(request, "manage/confirm_delete.html", {
        "object": attempt, "kind": "попытку",
        "warning": "Результат будет удалён безвозвратно, а попытка снова станет доступна пользователю.",
        "back_url": "manage:attempt_detail", "back_arg": pk,
    })
