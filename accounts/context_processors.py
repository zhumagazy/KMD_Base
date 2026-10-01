ADMIN_NAV = [
    ("Обзор", "manage:dashboard", "square.grid", ("dashboard",)),
    ("Группы", "manage:group_list", "folder", ("group",)),
    ("Курсы", "manage:course_list", "book", ("course", "material")),
    ("Тесты", "manage:test_list", "checklist", ("test", "question", "assign")),
    ("Результаты", "manage:results", "chart", ("result", "attempt")),
    ("Пользователи", "manage:user_list", "person.2", ("user",)),
]

EMPLOYEE_NAV = [
    ("Главная", "home", "house", ("home",)),
    ("Курсы", "courses", "book", ("course", "material")),
    ("Мои тесты", "my_tests", "checklist", ("my_tests", "test", "attempt")),
]


def navigation(request):
    user = getattr(request, "user", None)
    if not user or not user.is_authenticated:
        return {}
    match = getattr(request, "resolver_match", None)
    url_name = (match.url_name or "") if match else ""
    in_manage = bool(match and match.namespace == "manage")
    items = ADMIN_NAV if user.is_admin else EMPLOYEE_NAV
    nav = []
    for label, url, icon, keys in items:
        active = (in_manage == user.is_admin) and any(url_name.startswith(k) for k in keys)
        nav.append({"label": label, "url": url, "icon": icon, "active": active})
    return {"nav_items": nav}
