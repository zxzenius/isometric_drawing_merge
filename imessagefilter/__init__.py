# copy from:
# https://github.com/jondy/imessagefilter
import os
from ctypes import windll


class CMessageFilter(object):

    msgfilter = None

    @classmethod
    def register(cls):
        if cls.msgfilter is None:
            filename = 'msgfilter.dll'
            if not os.path.exists(filename):
                filename = os.path.join(os.path.dirname(__file__), filename)
            cls.msgfilter = windll.LoadLibrary(os.path.normpath(filename))
        cls.msgfilter.register_message_filter()

    @classmethod
    def revoke(cls):
        if cls.msgfilter:
            cls.msgfilter.revoke_message_filter()


if __name__ == '__main__':
    windll.ole32.CoInitialize(0)
    CMessageFilter.register()
    CMessageFilter.revoke()
    windll.ole32.CoUninitialize()