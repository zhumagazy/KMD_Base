from django.conf import settings
from django.contrib import admin
from django.contrib.auth import views as auth_views
from django.contrib.auth.decorators import login_required
from django.urls import include, path
from django.views.static import serve

from accounts import views as account_views
from exams import views as exam_views
from learning import views as learning_views

urlpatterns = [
    path("", account_views.home, name="home"),
    path("login/", auth_views.LoginView.as_view(template_name="accounts/login.html", redirect_authenticated_user=True), name="login"),
    path("logout/", auth_views.LogoutView.as_view(), name="logout"),
    path("password/", account_views.password_change, name="password_change"),
    # Сотрудник
    path("courses/", learning_views.course_catalog, name="courses"),
    path("courses/<int:pk>/", learning_views.course_detail, name="course_detail"),
    path("materials/<int:pk>/", learning_views.material_detail, name="material_detail"),
    path("materials/<int:pk>/complete/", learning_views.material_complete, name="material_complete"),
    path("tests/", exam_views.my_tests, name="my_tests"),
    path("tests/<int:pk>/", exam_views.test_intro, name="test_intro"),
    path("tests/<int:pk>/start/", exam_views.test_start, name="test_start"),
    path("attempts/<int:pk>/", exam_views.attempt_take, name="attempt_take"),
    path("attempts/<int:pk>/save/", exam_views.attempt_save, name="attempt_save"),
    path("attempts/<int:pk>/submit/", exam_views.attempt_submit, name="attempt_submit"),
    path("attempts/<int:pk>/done/", exam_views.attempt_done, name="attempt_done"),
    # Администратор
    path("manage/", include("config.manage_urls", namespace="manage")),
    path("django-admin/", admin.site.urls),
]

# Загруженные материалы доступны только авторизованным пользователям.
urlpatterns += [
    path(f"{settings.MEDIA_URL.strip('/')}/<path:path>", login_required(serve), {"document_root": settings.MEDIA_ROOT}),
]
