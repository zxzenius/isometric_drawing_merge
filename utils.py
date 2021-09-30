from pathlib import Path

import pythoncom
from win32com.client import VARIANT
from win32com.client.gencache import EnsureDispatch
from win32com.client.dynamic import Dispatch
from collections import namedtuple


def get_application(prog_id: str):
    # ref:https://gist.github.com/rdapaz/63590adb94a46039ca4a10994dff9dbe#gistcomment-2918299
    try:
        # return EnsureDispatch(prog_id)
        return Dispatch(prog_id)
    except AttributeError:
        import re
        import sys
        import shutil
        import win32com
        # Remove cache and try again.
        gen_path = win32com.__gen_path__
        modules = [m.__name__ for m in sys.modules.values()]
        for module in modules:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        # Remove gen_py folder
        shutil.rmtree(gen_path)
        # reload
        # return win32com.client.gencache.EnsureDispatch(prog_id)
        return win32com.client.dynamic.Dispatch(prog_id)


def get_acad():
    return get_application("AutoCAD.Application")


def vt_point(x, y, z):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


Point = namedtuple("Point", ['x', 'y', 'z'])


def get_drawings(folder):
    """
    Get all of autocad dwg files
    :param folder:
    :return:
    List[Path] including all dwg files
    """
    # case insensitive in windows system, so "dwg" is ok
    return sorted(Path(folder).glob('**/*.dwg'))


def get_document(app, filename=None):
    # load current file
    if filename is None:
        return app.ActiveDocument

    for document in app.Documents:
        if Path(document.FullName) == Path(filename):
            return document

    return app.Documents.Open(filename)


separator = '='*20
sub_separator = '*' * 10
