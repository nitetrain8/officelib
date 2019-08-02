import win32com.client as w32_client
import win32com.client.gencache as w32c_gen
from pywintypes import com_error  # pylint: disable=I0011,E0611
from win32com.client import constants as wincom_const
from officelib import const as olconst
import contextlib

class WdC():
    def __getattr__(self, a):
        v = getattr(wincom_const, a, None)
        if v is None:
            try:
                v = getattr(olconst, a)
            except AttributeError:
                raise AttributeError("'%s' not found."%a) from None
        else:
            assert v == getattr(olconst, a)
        object.__setattr__(self, a, v)
        return v

wdc = WdC()

_unspecified = object()

@contextlib.contextmanager
def screen_lock(word, visible=_unspecified):
    word.ScreenUpdating = False
    if visible is not _unspecified:
        oldvis = word.Visible
        word.Visible = visible
    try:
        yield None
    finally:
        word.ScreenUpdating = True
        if visible is not _unspecified:
            word.Visible = oldvis
lock_screen = screen_lock


def Word(visible=True):
    try:
        w = w32_client.GetObject("Word.Application")
    except com_error:
        w = w32c_gen.EnsureDispatch("Word.Application")
    w.Visible = visible
    return w


def format_list_level(ll, level):
    s = "".join("%%%d." % j for j in range(1, level+1))
    npos = (level-1)*0.25
    tpos = 0.25 + (level - 1) * 0.25

    ll.NumberFormat = s
    ll.TrailingCharacter = c.wdTrailingTab
    ll.NumberStyle = c.wdListNumberStyleArabic
    ll.NumberPosition = inches_to_points(npos)
    ll.Alignment = c.wdListLevelAlignLeft
    ll.TextPosition = inches_to_points(tpos)
    ll.TabPosition = c.wdUndefined
    ll.ResetOnHigher = level - 1
    ll.StartAt = 1
    
    font = ll.Font
    font.Bold = c.wdUndefined
    font.Italic = c.wdUndefined
    font.StrikeThrough = c.wdUndefined
    font.Subscript = c.wdUndefined
    font.Superscript = c.wdUndefined
    font.Shadow = c.wdUndefined
    font.Outline = c.wdUndefined
    font.Emboss = c.wdUndefined
    font.Engrave = c.wdUndefined
    font.AllCaps = c.wdUndefined
    font.Hidden = c.wdUndefined
    font.Underline = c.wdUndefined
    font.Color = c.wdUndefined
    font.Size = c.wdUndefined
    font.Animation = c.wdUndefined
    font.DoubleStrikeThrough = c.wdUndefined
    font.Name = ""


def make_list(word):
    gal = word.ListGalleries(c.wdOutlineNumberGallery)
    lt = gal.ListTemplates(1)
    for level in range(1, 7):
        ll = lt.ListLevels(level)
        format_list_level(ll, level)
    lt.Name = ""
    return lt


def apply_list_template(lt, rng):
    rng.ListFormat.ApplyListTemplateWithLevel(ListTemplate=lt,
            ContinuePreviousList=False, ApplyTo=c.wdListApplyToWholeList,
            DefaultListBehavior=c.wdWord10ListBehavior)


def convert(factor, from_unit, to_unit):
    src = \
"""
def %s_to_%s(%s):
    return %s * %s
""" % (from_unit, to_unit, from_unit, from_unit, factor)
    ns = {}
    exec(src, ns, ns)
    return ns['%s_to_%s' % (from_unit, to_unit)]

inches_to_points = convert(72, 'inches', 'points')
points_to_inches = convert(1/72, 'points', 'inches')
