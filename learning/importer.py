"""Импорт курсов и тестов из папки.

Структура:
    content/
      01 Адаптация/                 ← группа курсов (папка с подпапками)
        описание.txt                ← необязательное описание группы
        01 Знакомство с компанией/  ← курс
          описание.txt              ← описание курса
          01 О компании.docx        ← материалы: .docx .md .txt .html .pdf .pptx .xlsx .mp4 .url …
          02 Видео.url              ← ярлык на ссылку (YouTube/Vimeo станет встроенным видео)
          тест.txt                  ← тест курса (формат — в README)
      Отдельный курс/               ← папка только с файлами = курс без группы
      Тесты/                        ← тесты без курса (например, для кандидатов)
        Входной тест.txt

Числа в начале имён задают порядок и в название не попадают. Повторный запуск
обновляет существующие записи по названию, а не создаёт дубликаты.
"""
import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path

from django.core.files import File
from django.core.files.base import ContentFile
from django.core.files.storage import default_storage
from django.db import transaction
from django.utils.text import get_valid_filename

from exams.models import Choice, Question, Test

from .models import Course, CourseGroup, Material

ORDER_RE = re.compile(r"^\s*(\d+)[\s._)\-]+(.+)$")
SKIP_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}
DESCRIPTION_NAMES = {"описание", "description", "readme"}
TEST_NAME_RE = re.compile(r"^(тест|test)\b", re.IGNORECASE)
TEST_DIR_NAMES = {"тесты", "tests"}
TEXT_EXT = {".md", ".txt"}
VIDEO_EXT = {".mp4", ".webm", ".mov", ".m4v"}
YES = {"да", "yes", "true", "1", "+"}


@dataclass
class Report:
    created: list = field(default_factory=list)
    updated: list = field(default_factory=list)
    skipped: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    unchanged: int = 0


def split_order(name):
    m = ORDER_RE.match(name)
    if m:
        return int(m.group(1)), m.group(2).strip()
    return 0, name.strip()


def read_text(path):
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def file_hash(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def visible(path):
    n = path.name
    return not (n.startswith(".") or n.startswith("~$") or n.lower() in SKIP_NAMES)


def is_description(path):
    return path.is_file() and path.stem.lower() in DESCRIPTION_NAMES and path.suffix.lower() in TEXT_EXT


def is_test_file(path):
    return path.is_file() and path.suffix.lower() in TEXT_EXT and bool(TEST_NAME_RE.match(split_order(path.stem)[1]))


class Importer:
    def __init__(self, root, dry_run=False, stdout=None):
        self.root = Path(root)
        self.dry_run = dry_run
        self.report = Report()
        self.out = stdout

    def log(self, msg):
        if self.out:
            self.out.write(msg)

    # ---------- обход папок ----------

    def run(self):
        if not self.root.is_dir():
            raise FileNotFoundError(f"Папка {self.root} не найдена")
        with transaction.atomic():
            for top in sorted(p for p in self.root.iterdir() if visible(p)):
                if top.is_file():
                    if is_test_file(top):
                        self.import_test(top, course=None)
                    continue
                if top.name.lower() in TEST_DIR_NAMES:
                    for f in sorted(top.iterdir()):
                        if visible(f) and f.is_file() and f.suffix.lower() in TEXT_EXT:
                            self.import_test(f, course=None, title_from_file=True)
                    continue
                subdirs = [p for p in top.iterdir() if p.is_dir() and visible(p)]
                if subdirs:
                    self.import_group(top, subdirs)
                else:
                    self.import_course(top, group=None)
            if self.dry_run:
                transaction.set_rollback(True)
        return self.report

    def import_group(self, folder, subdirs):
        order, title = split_order(folder.name)
        group, created = CourseGroup.objects.get_or_create(title=title, defaults={"order": order})
        desc = next((p for p in folder.iterdir() if is_description(p)), None)
        group.order = order or group.order
        if desc:
            group.description = read_text(desc).strip()
        group.save()
        if created:
            self.report.created.append(f"Группа «{title}»")
        for sub in sorted(subdirs):
            self.import_course(sub, group=group)

    def import_course(self, folder, group):
        order, title = split_order(folder.name)
        course, created = Course.objects.get_or_create(title=title, group=group, defaults={"order": order})
        course.order = order or course.order
        if created:
            self.report.created.append(f"Курс «{title}»")
        files = sorted(p for p in folder.rglob("*") if p.is_file() and visible(p) and all(visible(x) for x in p.relative_to(folder).parents if str(x) != "."))
        for f in files:
            try:
                if is_description(f) and f.parent == folder:
                    course.description = read_text(f).strip()
                elif is_test_file(f):
                    self.import_test(f, course=course)
                else:
                    self.import_material(f, course, folder)
            except Exception as e:  # noqa: BLE001 — один битый файл не должен ронять весь импорт
                self.report.errors.append(f"{f.relative_to(self.root)}: {e}")
        course.save()

    # ---------- материалы ----------

    def import_material(self, path, course, course_dir):
        rel = path.relative_to(course_dir)
        order, title = split_order(path.stem)
        if len(rel.parts) > 1:
            # Файлы во вложенных папках: «Раздел / Файл»
            title = " / ".join([split_order(p)[1] for p in rel.parts[:-1]] + [title])
        ext = path.suffix.lower()
        digest = file_hash(path)
        m = Material.objects.filter(course=course, title=title).first()
        if m and m.source_hash == digest:
            self.report.unchanged += 1  # файл не менялся — правки, сделанные на сайте, сохраняются
            return
        created = m is None
        if created:
            m = Material(course=course, title=title)
        m.source_hash = digest
        m.order = order or m.order or (course.materials.count() + 1)
        m.body, m.url = "", ""

        if ext in TEXT_EXT:
            m.kind = Material.Kind.TEXT
            m.body = read_text(path)
        elif ext in (".html", ".htm"):
            m.kind = Material.Kind.TEXT
            m.body = read_text(path)
        elif ext == ".docx":
            m.kind = Material.Kind.TEXT
            m.body = self.docx_to_html(path, course)
        elif ext == ".url":
            url = self.parse_url_file(path)
            if not url:
                self.report.skipped.append(f"{rel}: в ярлыке нет ссылки")
                return
            m.url = url
            m.kind = Material.Kind.VIDEO if any(s in url for s in ("youtube.", "youtu.be", "vimeo.")) else Material.Kind.LINK
        else:
            m.kind = Material.Kind.VIDEO if ext in VIDEO_EXT else Material.Kind.FILE
            self.attach_file(m, path, course)
        if ext in TEXT_EXT or ext in (".html", ".htm", ".docx", ".url"):
            m.file = None
        m.save()
        (self.report.created if created else self.report.updated).append(f"  материал «{title}»")

    def attach_file(self, material, path, course):
        name = f"materials/import/{course.pk}/{get_valid_filename(path.name)}"
        if default_storage.exists(name) and default_storage.size(name) == path.stat().st_size:
            material.file.name = name
            return
        if self.dry_run:
            material.file.name = name
            return
        if default_storage.exists(name):
            default_storage.delete(name)
        with path.open("rb") as fh:
            material.file.name = default_storage.save(name, File(fh))

    def docx_to_html(self, path, course):
        import mammoth

        counter = {"n": 0}

        def save_image(image):
            counter["n"] += 1
            ext = (image.content_type or "image/png").split("/")[-1].replace("jpeg", "jpg")
            name = f"materials/import/{course.pk}/img/{get_valid_filename(path.stem)}-{counter['n']}.{ext}"
            if not self.dry_run:
                if default_storage.exists(name):
                    default_storage.delete(name)
                with image.open() as fh:
                    name = default_storage.save(name, ContentFile(fh.read()))
            return {"src": default_storage.url(name)}

        with path.open("rb") as fh:
            result = mammoth.convert_to_html(fh, convert_image=mammoth.images.img_element(save_image))
        return result.value

    @staticmethod
    def parse_url_file(path):
        for line in read_text(path).splitlines():
            if line.strip().upper().startswith("URL="):
                return line.split("=", 1)[1].strip()
        text = read_text(path).strip()
        return text if text.startswith("http") else ""

    # ---------- тесты ----------

    def import_test(self, path, course, title_from_file=False):
        digest = file_hash(path)
        spec = parse_test(read_text(path))
        title = spec["title"]
        if not title:
            stem_title = split_order(path.stem)[1]
            title = stem_title if title_from_file or not course else f"Тест: {course.title}"
        test = Test.objects.filter(title=title, course=course).first()
        if test and test.source_hash == digest:
            self.report.unchanged += 1
            return
        if test and test.attempts.exists():
            self.report.skipped.append(f"Тест «{title}»: уже есть попытки, вопросы не изменены")
            return
        created = test is None
        if created:
            test = Test(title=title, course=course)
        for key, value in spec["settings"].items():
            setattr(test, key, value)
        test.source_hash = digest
        test.save()
        test.questions.all().delete()
        for i, q in enumerate(spec["questions"], start=1):
            n_correct = sum(1 for _, ok in q["choices"] if ok)
            if not q["choices"]:
                kind = Question.Kind.TEXT
            elif n_correct > 1:
                kind = Question.Kind.MULTIPLE
            else:
                kind = Question.Kind.SINGLE
            if q["choices"] and n_correct == 0:
                self.report.errors.append(f"{path.name}: в вопросе «{q['text'][:40]}» не отмечен верный ответ (+)")
            question = Question.objects.create(
                test=test, order=i, kind=kind, text=q["text"], points=q["points"], explanation=q["explanation"]
            )
            Choice.objects.bulk_create(
                [Choice(question=question, text=t[:500], is_correct=ok, order=j) for j, (t, ok) in enumerate(q["choices"], 1)]
            )
        (self.report.created if created else self.report.updated).append(
            f"  тест «{title}» ({len(spec['questions'])} вопр.)"
        )


SETTINGS_MAP = {
    "время": ("time_limit_minutes", int),
    "проходной балл": ("pass_percent", int),
    "проходной": ("pass_percent", int),
    "вопросов в попытке": ("questions_per_attempt", int),
    "показывать результат": ("show_results", lambda v: v.lower() in YES),
    "показывать ответы": ("show_correct_answers", lambda v: v.lower() in YES),
    "перемешивать": ("shuffle_questions", lambda v: v.lower() in YES),
    "после курса": ("require_course_completion", lambda v: v.lower() in YES),
    "описание": ("description", str),
}


def parse_test(text):
    """Простой текстовый формат теста:

        # Название теста
        Время: 20
        Проходной балл: 70

        ? Текст вопроса
        + верный вариант
        - неверный вариант
        Баллы: 2
        Пояснение: почему так

        ? Вопрос без вариантов — развёрнутый ответ
    """
    spec = {"title": "", "settings": {}, "questions": []}
    current = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("#") and not spec["title"] and not spec["questions"]:
            spec["title"] = line.lstrip("#").strip()
            continue
        if line.startswith("?"):
            current = {"text": line[1:].strip(), "choices": [], "points": 1, "explanation": ""}
            spec["questions"].append(current)
            continue
        if current and line[0] in "+-" and len(line) > 1:
            current["choices"].append((line[1:].strip(), line[0] == "+"))
            continue
        if current and line[0] in "*" and len(line) > 1:
            current["choices"].append((line[1:].strip(), True))
            continue
        if ":" in line:
            key, value = (s.strip() for s in line.split(":", 1))
            k = key.lower()
            if current and k == "баллы" and value.isdigit():
                current["points"] = max(int(value), 1)
                continue
            if current and k == "пояснение":
                current["explanation"] = value
                continue
            if not current and k in SETTINGS_MAP:
                attr, conv = SETTINGS_MAP[k]
                try:
                    spec["settings"][attr] = conv(value.replace("%", "").replace("мин", "").strip())
                except ValueError:
                    pass
                continue
        if current and not current["choices"]:
            current["text"] += "\n" + line
    return spec
