# depenses/templatetags/custom_filters.py
from django import template

register = template.Library()

@register.filter
def split(value, arg):
    """Sépare une chaîne par un séparateur"""
    return value.split(arg)