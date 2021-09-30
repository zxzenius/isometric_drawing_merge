from pathlib import Path
from pywintypes import com_error
import os

from utils import get_acad, vt_point, get_drawings, separator
from imessagefilter import CMessageFilter

A3_width = 420
A3_height = 297


def merge(folder):
    drawings = get_drawings(folder)
    if not drawings:
        print('No dwg file in "%s"' % Path(folder).absolute())
        return
    CMessageFilter.register()
    app = get_acad()
    doc = app.Documents.Add()
    model_space = doc.ModelSpace

    counter = 0
    x = 0
    y = 0
    z = 0
    scale_x = 1
    scale_y = 1
    scale_z = 1
    rotation = 0
    offset_x = A3_width + 50
    offset_y = -A3_height - 50

    print('Merging...')
    print(separator)

    for dwg_file in drawings:
        print(dwg_file)
        try:
            model_space.InsertBlock(vt_point(x, y, z), dwg_file, scale_x, scale_y, scale_z, rotation)
            print('Inserted')
            y += offset_y
            counter += 1
        except com_error as e:
            print('Error (%s), skipped' % e.excepinfo[2])

    CMessageFilter.revoke()
    print(separator)
    print('finished, %s/%s files merged.' % (counter, len(drawings)))


if __name__ == '__main__':
    merge('')
    # input('Press any key to exit')
    os.system('pause')
