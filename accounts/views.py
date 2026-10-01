from django.contrib import messages
from django.contrib.auth import update_session_auth_hash
from django.contrib.auth.decorators import login_required
from django.contrib.auth.forms import PasswordChangeForm
from django.db.models import Count, Q
from django.shortcuts import get_object_or_404, redirect, render
from django.views.decorators.http import require_POST

from exams.models import Assignment, Attempt, Test
from learning.models import Course, CourseAccess, CourseGroup, GroupAccess

from .decorators import admin_required
from .forms import AdminSetPasswordForm, UserCreateForm, UserEditForm
from .models import User


@login_required
def home(request):
    if request.user.is_admin:
        return redirect("manage:dashboard")
    from learning.views import employee_home

    return employee_home(request)


@login_required
def password_change(request):
    form = PasswordChangeForm(request.user, request.POST or None)
    if request.method == "POST" and form.is_valid():
        user = form.save()
        update_session_auth_hash(request, user)
        messages.success(request, "Пароль изменён.")
        return redirect("home")
    return render(request, "accounts/password_change.html", {"form": form})


# ---------- Администрирование пользователей ----------


@admin_required
def user_list(request):
    users = User.objects.all()
    q = request.GET.get("q", "").strip()
    kind = request.GET.get("kind", "")
    status = request.GET.get("status", "active")
    if q:
        users = users.filter(
            Q(username__icontains=q) | Q(first_name__icontains=q) | Q(last_name__icontains=q)
            | Q(email__icontains=q) | Q(position__icontains=q)
        )
    if kind == "employee":
        users = users.filter(role=User.Role.EMPLOYEE, is_candidate=False)
    elif kind == "candidate":
        users = users.filter(is_candidate=True)
    elif kind == "admin":
        users = users.filter(Q(role=User.Role.ADMIN) | Q(is_superuser=True))
    if status == "active":
        users = users.filter(is_active=True)
    elif status == "inactive":
        users = users.filter(is_active=False)
    users = users.annotate(
        n_groups=Count("group_accesses", distinct=True),
        n_courses=Count("course_accesses", distinct=True),
        n_tests=Count("test_assignments", distinct=True),
    )
    return render(request, "manage/user_list.html", {"users": users, "q": q, "kind": kind, "status": status})


@admin_required
def user_create(request):
    form = UserCreateForm(request.POST or None)
    if request.method == "POST" and form.is_valid():
        user = form.save()
        messages.success(request, f"Пользователь «{user}» создан. Теперь выдайте ему доступы.")
        return redirect("manage:user_access", user.pk)
    return render(request, "manage/form.html", {
        "form": form, "title": "Новый пользователь", "back_url": "manage:user_list", "submit": "Создать",
    })


@admin_required
def user_detail(request, pk):
    u = get_object_or_404(User, pk=pk)
    return render(request, "manage/user_detail.html", {
        "u": u,
        "group_accesses": u.group_accesses.select_related("group"),
        "course_accesses": u.course_accesses.select_related("course", "course__group"),
        "assignments": u.test_assignments.select_related("test").prefetch_related("attempts"),
        "attempts": u.attempts.select_related("test").exclude(status=Attempt.Status.IN_PROGRESS)[:20],
    })


@admin_required
def user_edit(request, pk):
    u = get_object_or_404(User, pk=pk)
    form = UserEditForm(request.POST or None, instance=u, editor=request.user)
    if request.method == "POST" and form.is_valid():
        form.save()
        messages.success(request, "Изменения сохранены.")
        return redirect("manage:user_detail", u.pk)
    return render(request, "manage/form.html", {
        "form": form, "title": f"Редактирование: {u}", "back_url": "manage:user_detail", "back_arg": u.pk,
    })


@admin_required
def user_password(request, pk):
    u = get_object_or_404(User, pk=pk)
    form = AdminSetPasswordForm(u, request.POST or None)
    if request.method == "POST" and form.is_valid():
        form.save()
        if u.pk == request.user.pk:
            update_session_auth_hash(request, u)
        messages.success(request, f"Пароль пользователя «{u}» изменён.")
        return redirect("manage:user_detail", u.pk)
    return render(request, "manage/form.html", {
        "form": form, "title": f"Новый пароль: {u}", "back_url": "manage:user_detail", "back_arg": u.pk,
    })


@admin_required
@require_POST
def user_toggle_active(request, pk):
    u = get_object_or_404(User, pk=pk)
    if u.pk == request.user.pk:
        messages.error(request, "Нельзя заблокировать собственную учётную запись.")
    else:
        u.is_active = not u.is_active
        u.save(update_fields=["is_active"])
        messages.success(request, "Пользователь активирован." if u.is_active else "Пользователь заблокирован.")
    return redirect("manage:user_detail", u.pk)


@admin_required
def user_delete(request, pk):
    u = get_object_or_404(User, pk=pk)
    if u.pk == request.user.pk:
        messages.error(request, "Нельзя удалить собственную учётную запись.")
        return redirect("manage:user_detail", u.pk)
    if request.method == "POST":
        name = str(u)
        u.delete()
        messages.success(request, f"Пользователь «{name}» удалён.")
        return redirect("manage:user_list")
    return render(request, "manage/confirm_delete.html", {
        "object": u, "kind": "пользователя",
        "warning": "Будут удалены все его доступы, попытки и результаты тестов.",
        "back_url": "manage:user_detail", "back_arg": u.pk,
    })


@admin_required
def user_access(request, pk):
    """Доступы пользователя: группы курсов, отдельные курсы и тесты на одной странице."""
    u = get_object_or_404(User, pk=pk)
    groups = CourseGroup.objects.prefetch_related("courses")
    standalone = Course.objects.filter(group__isnull=True)
    tests = Test.objects.filter(is_active=True)
    if request.method == "POST":
        group_ids = {int(x) for x in request.POST.getlist("groups")}
        course_ids = {int(x) for x in request.POST.getlist("courses")}
        test_ids = {int(x) for x in request.POST.getlist("tests")}
        GroupAccess.objects.filter(user=u).exclude(group_id__in=group_ids).delete()
        CourseAccess.objects.filter(user=u).exclude(course_id__in=course_ids).delete()
        for gid in group_ids:
            GroupAccess.objects.get_or_create(user=u, group_id=gid, defaults={"granted_by": request.user})
        for cid in course_ids:
            CourseAccess.objects.get_or_create(user=u, course_id=cid, defaults={"granted_by": request.user})
        open_tests = {
            a.test_id for a in Assignment.objects.filter(user=u, is_revoked=False)
        }
        for tid in test_ids - open_tests:
            Assignment.objects.create(user=u, test_id=tid, created_by=request.user)
        messages.success(request, "Доступы обновлены.")
        return redirect("manage:user_detail", u.pk)
    return render(request, "manage/user_access.html", {
        "u": u,
        "groups": groups,
        "standalone": standalone,
        "tests": tests,
        "group_ids": set(u.group_accesses.values_list("group_id", flat=True)),
        "course_ids": set(u.course_accesses.values_list("course_id", flat=True)),
        "test_ids": set(u.test_assignments.filter(is_revoked=False).values_list("test_id", flat=True)),
    })
