from django.urls import path

from accounts import views as a
from exams import views as e
from learning import views as l

app_name = "manage"

urlpatterns = [
    path("", l.dashboard, name="dashboard"),
    # Группы
    path("groups/", l.group_list, name="group_list"),
    path("groups/new/", l.group_form, name="group_create"),
    path("groups/<int:pk>/", l.group_detail, name="group_detail"),
    path("groups/<int:pk>/edit/", l.group_form, name="group_edit"),
    path("groups/<int:pk>/delete/", l.group_delete, name="group_delete"),
    path("groups/<int:pk>/access/", l.group_access, name="group_access"),
    # Курсы и материалы
    path("courses/", l.course_list, name="course_list"),
    path("courses/new/", l.course_form, name="course_create"),
    path("courses/<int:pk>/", l.course_detail_admin, name="course_detail"),
    path("courses/<int:pk>/edit/", l.course_form, name="course_edit"),
    path("courses/<int:pk>/delete/", l.course_delete, name="course_delete"),
    path("courses/<int:pk>/access/", l.course_access, name="course_access"),
    path("courses/<int:course_pk>/materials/new/", l.material_form, name="material_create"),
    path("materials/<int:pk>/edit/", l.material_form, name="material_edit"),
    path("materials/<int:pk>/delete/", l.material_delete, name="material_delete"),
    # Тесты
    path("tests/", e.test_list, name="test_list"),
    path("tests/new/", e.test_form, name="test_create"),
    path("tests/<int:pk>/", e.test_detail, name="test_detail"),
    path("tests/<int:pk>/edit/", e.test_form, name="test_edit"),
    path("tests/<int:pk>/delete/", e.test_delete, name="test_delete"),
    path("tests/<int:pk>/assign/", e.assign, name="assign"),
    path("tests/<int:test_pk>/questions/new/", e.question_form, name="question_create"),
    path("questions/<int:pk>/edit/", e.question_form, name="question_edit"),
    path("questions/<int:pk>/delete/", e.question_delete, name="question_delete"),
    path("assignments/<int:pk>/edit/", e.assignment_edit, name="assign_edit"),
    path("assignments/<int:pk>/extra/", e.assignment_extra_attempt, name="assign_extra"),
    path("assignments/<int:pk>/revoke/", e.assignment_revoke, name="assign_revoke"),
    # Результаты
    path("results/", e.results, name="results"),
    path("results/export/", e.results_export, name="results_export"),
    path("attempts/<int:pk>/", e.attempt_detail, name="attempt_detail"),
    path("attempts/<int:pk>/delete/", e.attempt_delete, name="attempt_delete"),
    # Пользователи
    path("users/", a.user_list, name="user_list"),
    path("users/new/", a.user_create, name="user_create"),
    path("users/<int:pk>/", a.user_detail, name="user_detail"),
    path("users/<int:pk>/edit/", a.user_edit, name="user_edit"),
    path("users/<int:pk>/password/", a.user_password, name="user_password"),
    path("users/<int:pk>/toggle/", a.user_toggle_active, name="user_toggle"),
    path("users/<int:pk>/delete/", a.user_delete, name="user_delete"),
    path("users/<int:pk>/access/", a.user_access, name="user_access"),
]
