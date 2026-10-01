import markdown as md
from django.conf import settings
from django.db import models
from django.db.models import Q
from django.utils import timezone
from django.utils.safestring import mark_safe


def active_q(prefix=""):
    now = timezone.now()
    return Q(**{f"{prefix}expires_at__isnull": True}) | Q(**{f"{prefix}expires_at__gt": now})


class CourseGroup(models.Model):
    title = models.CharField("Название", max_length=200)
    description = models.TextField("Описание", blank=True)
    order = models.PositiveIntegerField("Порядок", default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Группа курсов"
        verbose_name_plural = "Группы курсов"
        ordering = ["order", "title"]

    def __str__(self):
        return self.title


class Course(models.Model):
    group = models.ForeignKey(
        CourseGroup, verbose_name="Группа", related_name="courses",
        on_delete=models.SET_NULL, null=True, blank=True,
    )
    title = models.CharField("Название", max_length=200)
    description = models.TextField("Описание", blank=True)
    is_published = models.BooleanField(
        "Опубликован", default=True, help_text="Неопубликованный курс не виден сотрудникам даже при наличии доступа."
    )
    order = models.PositiveIntegerField("Порядок", default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Курс"
        verbose_name_plural = "Курсы"
        ordering = ["order", "title"]

    def __str__(self):
        return self.title

    def progress_for(self, user):
        total = self.materials.count()
        if not total:
            return {"done": 0, "total": 0, "percent": 0}
        done = MaterialProgress.objects.filter(user=user, material__course=self).count()
        return {"done": done, "total": total, "percent": round(done * 100 / total)}

    def is_completed_by(self, user):
        p = self.progress_for(user)
        return p["total"] > 0 and p["done"] >= p["total"]


class Material(models.Model):
    class Kind(models.TextChoices):
        TEXT = "text", "Статья"
        VIDEO = "video", "Видео"
        FILE = "file", "Файл"
        LINK = "link", "Ссылка"

    course = models.ForeignKey(Course, verbose_name="Курс", related_name="materials", on_delete=models.CASCADE)
    title = models.CharField("Название", max_length=200)
    kind = models.CharField("Тип", max_length=16, choices=Kind.choices, default=Kind.TEXT)
    body = models.TextField(
        "Содержание", blank=True,
        help_text="Поддерживается Markdown: **жирный**, *курсив*, # Заголовок, - список, [ссылка](https://…).",
    )
    file = models.FileField("Файл", upload_to="materials/%Y/%m/", blank=True)
    url = models.URLField("Ссылка / видео", blank=True, help_text="YouTube, Vimeo или любая другая ссылка.")
    duration_minutes = models.PositiveIntegerField("Время на изучение, мин", null=True, blank=True)
    order = models.PositiveIntegerField("Порядок", default=0)
    source_hash = models.CharField(max_length=64, blank=True, editable=False)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Материал"
        verbose_name_plural = "Материалы"
        ordering = ["order", "id"]

    def __str__(self):
        return self.title

    @property
    def filename(self):
        return self.file.name.rsplit("/", 1)[-1] if self.file else ""

    @property
    def extension(self):
        return ("." + self.filename.rsplit(".", 1)[-1].lower()) if "." in self.filename else ""

    @property
    def is_video_file(self):
        return self.extension in (".mp4", ".webm", ".mov", ".m4v")

    @property
    def is_pdf(self):
        return self.extension == ".pdf"

    @property
    def body_html(self):
        # Материалы пишут только администраторы, поэтому HTML внутри Markdown разрешён.
        return mark_safe(md.markdown(self.body, extensions=["extra", "sane_lists", "nl2br"]))

    @property
    def embed_url(self):
        url = self.url or ""
        if "youtube.com/watch" in url and "v=" in url:
            vid = url.split("v=", 1)[1].split("&", 1)[0]
            return f"https://www.youtube.com/embed/{vid}"
        if "youtu.be/" in url:
            vid = url.split("youtu.be/", 1)[1].split("?", 1)[0]
            return f"https://www.youtube.com/embed/{vid}"
        if "vimeo.com/" in url and "player.vimeo.com" not in url:
            vid = url.rstrip("/").rsplit("/", 1)[-1]
            if vid.isdigit():
                return f"https://player.vimeo.com/video/{vid}"
        if "player.vimeo.com" in url or "youtube.com/embed" in url:
            return url
        return ""


class AccessBase(models.Model):
    granted_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, related_name="+", on_delete=models.SET_NULL, null=True, blank=True
    )
    granted_at = models.DateTimeField(auto_now_add=True)
    expires_at = models.DateTimeField("Доступ до", null=True, blank=True)

    class Meta:
        abstract = True

    @property
    def is_active(self):
        return self.expires_at is None or self.expires_at > timezone.now()


class GroupAccess(AccessBase):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, related_name="group_accesses", on_delete=models.CASCADE)
    group = models.ForeignKey(CourseGroup, related_name="accesses", on_delete=models.CASCADE)

    class Meta:
        verbose_name = "Доступ к группе"
        verbose_name_plural = "Доступы к группам"
        constraints = [models.UniqueConstraint(fields=["user", "group"], name="uniq_group_access")]


class CourseAccess(AccessBase):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, related_name="course_accesses", on_delete=models.CASCADE)
    course = models.ForeignKey(Course, related_name="accesses", on_delete=models.CASCADE)

    class Meta:
        verbose_name = "Доступ к курсу"
        verbose_name_plural = "Доступы к курсам"
        constraints = [models.UniqueConstraint(fields=["user", "course"], name="uniq_course_access")]


class MaterialProgress(models.Model):
    user = models.ForeignKey(settings.AUTH_USER_MODEL, related_name="material_progress", on_delete=models.CASCADE)
    material = models.ForeignKey(Material, related_name="progress", on_delete=models.CASCADE)
    completed_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        constraints = [models.UniqueConstraint(fields=["user", "material"], name="uniq_material_progress")]


def accessible_courses(user):
    """Курсы, которые видит пользователь: через доступ к группе или к самому курсу."""
    if user.is_admin:
        return Course.objects.all()
    group_ids = GroupAccess.objects.filter(active_q(), user=user).values("group_id")
    course_ids = CourseAccess.objects.filter(active_q(), user=user).values("course_id")
    return Course.objects.filter(is_published=True).filter(Q(group_id__in=group_ids) | Q(id__in=course_ids))
