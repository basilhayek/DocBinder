import os
if os.name == "nt":
    from win32com.client import Dispatch

class AppHandler():
    ''' Base class for the class to handle getting path information and opening files '''


    def __init__(self):
        pass

    def getpath(self, doc):
        pass
        
    def openfile(self, path, filename):
        pass
    

class XLSHandler(AppHandler):

    def __init__(self):
        self._xl = Dispatch('Excel.Application')

    def getpath(self, doc):
        # Thanks to the example at https://gist.github.com/mikepsn/27dd0d768ccede849051
        return self._xl.Workbooks(doc).Path
        
    def openfile(self, path, filename):
        fqpath = '"{}/{}"'.format(path, filename)
        print(" [OPENING] {}".format(filename))
        self._xl.Workbooks.Open("{}".format(fqpath))


class PPTHandler(AppHandler):

    def __init__(self):
        self._pp = Dispatch('PowerPoint.Application')

    def getpath(self, doc):
        return self._pp.Presentations(doc).Path
        
    def openfile(self, path, filename):
        fqpath = '"{}/{}"'.format(path, filename)
        print(" [OPENING] {}".format(filename))
        self._pp.Presentations.Open("{}".format(fqpath))


class DOCHandler(AppHandler):

    def __init__(self):
        self._wd = Dispatch('Word.Application')

    def getpath(self, doc):
        return self._wd.Documents(doc).Path
        
    def openfile(self, path, filename):
        fqpath = '"{}/{}"'.format(path, filename)
        print(" [OPENING] {}".format(filename))
        self._wd.Documents.Open("{}".format(fqpath))

class MOCKHandler(AppHandler):

    def __init__(self):
        pass

    def getpath(self, doc):
        return "/mockpath"
        
    def openfile(self, path, filename):
        fqpath = '"{}/{}"'.format(path, filename)
        print(" [MOCK--OPENING] {}".format(filename))


class AppHandlerFactory():

    _appdict = {"Excel": XLSHandler,
                "PowerPoint": PPTHandler,
                "Word": DOCHandler
                }
    _appdictmock = {"Excel": MOCKHandler,
                "PowerPoint": MOCKHandler,
                "Word": MOCKHandler
                }

    def __init__(self, mock = False):
        if mock == True:
            self._appdict = self._appdictmock
        
    def applist(self):
        return self._appdict.keys()
        
    def gethandler(self, app):
        ''' Returns the class to handle a given file type '''
        return self._appdict[app]()
