from django import template
from django.utils.html import format_html

register = template.Library()


@register.simple_tag
def icon(name, cls=""):
    return format_html('<svg class="icon {}" aria-hidden="true"><use href="#i-{}"></use></svg>', cls, name)


@register.filter
def get(mapping, key):
    try:
        return mapping.get(key)
    except AttributeError:
        return None


@register.filter
def widget_type(bound_field):
    return bound_field.field.widget.__class__.__name__


@register.filter
def duration(td):
    if not td:
        return "—"
    total = int(td.total_seconds())
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    if h:
        return f"{h} ч {m} мин"
    if m:
        return f"{m} мин {s} с"
    return f"{s} с"


@register.filter
def plural(n, forms):
    """{{ n|plural:"курс,курса,курсов" }}"""
    one, few, many = forms.split(",")
    n = abs(int(n or 0))
    if n % 10 == 1 and n % 100 != 11:
        word = one
    elif 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        word = few
    else:
        word = many
    return f"{n} {word}"
