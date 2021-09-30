import os
from comtypes.client import CreateObject
from imessagefilter import CMessageFilter


def iter_excel_sheets(filename):
    CMessageFilter.register()
    aobj = CreateObject('Excel.Application', dynamic=True)
    aobj.Workbooks.Open(filename)
    for sheet in aobj.ActiveWorkbook.Worksheets:
        print('Processing %s' % sheet.Name)
    CMessageFilter.revoke()


if __name__ == '__main__':
    iter_excel_sheets(os.path.abspath('base2003.xls'))
