from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Count, Q
from django.http import Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_POST

from accounts.decorators import admin_required
from accounts.models import User
from exams.models import Assignment, Attempt, Test

from .forms import AccessExpiryForm, CourseForm, CourseGroupForm, MaterialForm
from .models import (
    Course, CourseAccess, CourseGroup, GroupAccess, Material, MaterialProgress, accessible_courses,
)


# ---------- Сотрудник ----------


def _grouped_courses(user):
    courses = list(accessible_courses(user).select_related("group").annotate(n_materials=Count("materials")))
    done = set(MaterialProgress.objects.filter(user=user).values_list("material__course_id", "material_id"))
    done_by_course = {}
    for cid, _ in done:
        done_by_course[cid] = done_by_course.get(cid, 0) + 1
    groups = {}
    for c in courses:
        c.n_done = done_by_course.get(c.id, 0)
        c.percent = round(c.n_done * 100 / c.n_materials) if c.n_materials else 0
        key = c.group
        groups.setdefault(key, []).append(c)
    ordered = sorted(groups.items(), key=lambda kv: (kv[0] is None, kv[0].order if kv[0] else 0, str(kv[0] or "")))
    return [{"group": g, "courses": cs} for g, cs in ordered], courses


def employee_home(request):
    grouped, courses = _grouped_courses(request.user)
    assignments = [
        a for a in Assignment.objects.filter(user=request.user, is_revoked=False, test__is_active=True)
        .select_related("test", "test__course")
        if a.state() in ("available", "in_progress", "scheduled", "locked")
    ]
    in_progress = [c for c in courses if 0 < c.percent < 100]
    return render(request, "employee/home.html", {
        "grouped": grouped, "courses": courses, "assignments": assignments, "in_progress": in_progress,
    })


@login_required
def course_catalog(request):
    grouped, courses = _grouped_courses(request.user)
    return render(request, "employee/courses.html", {"grouped": grouped, "courses": courses})


def _get_course_for(user, pk):
    try:
        return accessible_courses(user).select_related("group").get(pk=pk)
    except Course.DoesNotExist:
        raise Http404


@login_required
def course_detail(request, pk):
    course = _get_course_for(request.user, pk)
    materials = list(course.materials.all())
    done = set(MaterialProgress.objects.filter(user=request.user, material__course=course)
               .values_list("material_id", flat=True))
    for m in materials:
        m.done = m.id in done
    assignments = Assignment.objects.filter(
        user=request.user, test__course=course, is_revoked=False, test__is_active=True
    ).select_related("test")
    return render(request, "employee/course_detail.html", {
        "course": course, "materials": materials, "progress": course.progress_for(request.user),
        "assignments": assignments,
    })


@login_required
def material_detail(request, pk):
    material = get_object_or_404(Material.objects.select_related("course"), pk=pk)
    course = _get_course_for(request.user, material.course_id)
    siblings = list(course.materials.all())
    idx = next(i for i, m in enumerate(siblings) if m.pk == material.pk)
    done = MaterialProgress.objects.filter(user=request.user, material=material).exists()
    return render(request, "employee/material_detail.html", {
        "course": course, "material": material, "done": done,
        "prev": siblings[idx - 1] if idx > 0 else None,
        "next": siblings[idx + 1] if idx + 1 < len(siblings) else None,
        "position": idx + 1, "total": len(siblings),
    })


@login_required
@require_POST
def material_complete(request, pk):
    material = get_object_or_404(Material, pk=pk)
    _get_course_for(request.user, material.course_id)
    if request.POST.get("undo"):
        MaterialProgress.objects.filter(user=request.user, material=material).delete()
    else:
        MaterialProgress.objects.get_or_create(user=request.user, material=material)
    nxt = request.POST.get("next")
    if nxt and nxt.isdigit():
        return redirect("material_detail", int(nxt))
    return redirect("material_detail", material.pk)


# ---------- Администратор: обзор ----------


@admin_required
def dashboard(request):
    active_users = User.objects.filter(is_active=True)
    recent = Attempt.objects.exclude(status=Attempt.Status.IN_PROGRESS).select_related("user", "test")[:8]
    return render(request, "manage/dashboard.html", {
        "stats": {
            "employees": active_users.filter(role=User.Role.EMPLOYEE, is_candidate=False).count(),
            "candidates": active_users.filter(is_candidate=True).count(),
            "groups": CourseGroup.objects.count(),
            "courses": Course.objects.count(),
            "materials": Material.objects.count(),
            "tests": Test.objects.filter(is_active=True).count(),
            "pending_review": Attempt.objects.filter(needs_review=True).count(),
            "in_progress": Attempt.objects.filter(status=Attempt.Status.IN_PROGRESS).count(),
        },
        "recent": recent,
    })


# ---------- Администратор: группы ----------


@admin_required
def group_list(request):
    groups = CourseGroup.objects.annotate(
        n_courses=Count("courses", distinct=True), n_users=Count("accesses", distinct=True)
    )
    return render(request, "manage/group_list.html", {"groups": groups})


@admin_required
def group_detail(request, pk):
    group = get_object_or_404(CourseGroup, pk=pk)
    courses = group.courses.annotate(n_materials=Count("materials", distinct=True), n_tests=Count("tests", distinct=True))
    accesses = group.accesses.select_related("user")
    return render(request, "manage/group_detail.html", {"group": group, "courses": courses, "accesses": accesses})


@admin_required
def group_form(request, pk=None):
    group = get_object_or_404(CourseGroup, pk=pk) if pk else None
    form = CourseGroupForm(request.POST or None, instance=group)
    if request.method == "POST" and form.is_valid():
        group = form.save()
        messages.success(request, "Группа сохранена.")
        return redirect("manage:group_detail", group.pk)
    return render(request, "manage/form.html", {
        "form": form, "title": "Редактирование группы" if pk else "Новая группа",
        "back_url": "manage:group_detail" if pk else "manage:group_list", "back_arg": pk,
    })


@admin_required
def group_delete(request, pk):
    group = get_object_or_404(CourseGroup, pk=pk)
    if request.method == "POST":
        group.delete()
        messages.success(request, "Группа удалена. Курсы из неё сохранены без группы.")
        return redirect("manage:group_list")
    return render(request, "manage/confirm_delete.html", {
        "object": group, "kind": "группу", "warning": "Курсы группы не удаляются — они останутся без группы.",
        "back_url": "manage:group_detail", "back_arg": pk,
    })


# ---------- Администратор: курсы ----------


@admin_required
def course_list(request):
    courses = Course.objects.select_related("group").annotate(
        n_materials=Count("materials", distinct=True), n_tests=Count("tests", distinct=True),
        n_users=Count("accesses", distinct=True),
    )
    return render(request, "manage/course_list.html", {"courses": courses})


@admin_required
def course_detail_admin(request, pk):
    course = get_object_or_404(Course.objects.select_related("group"), pk=pk)
    return render(request, "manage/course_detail.html", {
        "course": course,
        "materials": course.materials.all(),
        "tests": course.tests.annotate(n_questions=Count("questions")),
        "accesses": course.accesses.select_related("user"),
        "group_accesses": course.group.accesses.select_related("user") if course.group else [],
    })


@admin_required
def course_form(request, pk=None):
    course = get_object_or_404(Course, pk=pk) if pk else None
    initial = {}
    if not pk and request.GET.get("group"):
        initial["group"] = request.GET["group"]
    form = CourseForm(request.POST or None, instance=course, initial=initial)
    if request.method == "POST" and form.is_valid():
        course = form.save()
        messages.success(request, "Курс сохранён.")
        return redirect("manage:course_detail", course.pk)
    return render(request, "manage/form.html", {
        "form": form, "title": "Редактирование курса" if pk else "Новый курс",
        "back_url": "manage:course_detail" if pk else "manage:course_list", "back_arg": pk,
    })


@admin_required
def course_delete(request, pk):
    course = get_object_or_404(Course, pk=pk)
    if request.method == "POST":
        course.delete()
        messages.success(request, "Курс удалён.")
        return redirect("manage:course_list")
    return render(request, "manage/confirm_delete.html", {
        "object": course, "kind": "курс", "warning": "Будут удалены все материалы курса и прогресс сотрудников.",
        "back_url": "manage:course_detail", "back_arg": pk,
    })


# ---------- Администратор: материалы ----------


@admin_required
def material_form(request, course_pk=None, pk=None):
    material = get_object_or_404(Material, pk=pk) if pk else None
    course = material.course if material else get_object_or_404(Course, pk=course_pk)
    initial = {}
    if not material:
        last = course.materials.order_by("-order").first()
        initial["order"] = (last.order + 1) if last else 1
    form = MaterialForm(request.POST or None, request.FILES or None, instance=material, initial=initial)
    if request.method == "POST" and form.is_valid():
        m = form.save(commit=False)
        m.course = course
        m.save()
        messages.success(request, "Материал сохранён.")
        return redirect("manage:course_detail", course.pk)
    return render(request, "manage/form.html", {
        "form": form, "title": "Редактирование материала" if pk else f"Новый материал — {course}",
        "back_url": "manage:course_detail", "back_arg": course.pk, "multipart": True, "wide": True,
        "preview_url": material and material.pk,
    })


@admin_required
def material_delete(request, pk):
    material = get_object_or_404(Material, pk=pk)
    if request.method == "POST":
        course_pk = material.course_id
        material.delete()
        messages.success(request, "Материал удалён.")
        return redirect("manage:course_detail", course_pk)
    return render(request, "manage/confirm_delete.html", {
        "object": material, "kind": "материал", "back_url": "manage:course_detail", "back_arg": material.course_id,
    })


# ---------- Администратор: доступы к группе / курсу ----------


def _target_access(request, target, access_model, field, back_url):
    form = AccessExpiryForm(request.POST or None)
    users = User.objects.filter(is_active=True).exclude(role=User.Role.ADMIN).exclude(is_superuser=True)
    existing = {a.user_id: a for a in access_model.objects.filter(**{field: target})}
    if request.method == "POST" and form.is_valid():
        selected = {int(x) for x in request.POST.getlist("users")}
        access_model.objects.filter(**{field: target}).exclude(user_id__in=selected).delete()
        for uid in selected - set(existing):
            access_model.objects.create(
                user_id=uid, granted_by=request.user, expires_at=form.cleaned_data["expires_at"], **{field: target}
            )
        messages.success(request, "Доступы обновлены.")
        return redirect(back_url, target.pk)
    q = request.GET.get("q", "").strip()
    return render(request, "manage/target_access.html", {
        "target": target, "users": users, "existing": existing, "form": form,
        "back_url": back_url, "q": q,
        "kind": "группе курсов" if field == "group" else "курсу",
    })


@admin_required
def group_access(request, pk):
    return _target_access(request, get_object_or_404(CourseGroup, pk=pk), GroupAccess, "group", "manage:group_detail")


@admin_required
def course_access(request, pk):
    return _target_access(request, get_object_or_404(Course, pk=pk), CourseAccess, "course", "manage:course_detail")
