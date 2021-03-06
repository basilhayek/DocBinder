import os
if os.name == "nt":
    import win32gui
    import win32com.client

import re       # For parsing filenames
import pprint   # For saving to file

from apphandler import AppHandlerFactory

class DocBinder():
    
    _doclist = {}
    _docindex = []
    _winlist = []
    _winlistmock = []
    _mock = False
    _workspaces = {}

    def __init__(self, mock = False):
        self._mock = mock
        self._openbinders()

        self._ahf = AppHandlerFactory(mock)
        if(self._mock):
            self._winlistmock = ['210426%20Clilent%20Dashboard%20Mock-up%20draft (version 1).xlsb  -  AutoRecovered - Excel',
                        'Weekly Executive Report - Consolidated Numbers.xlsx - Excel',
                        '2021 Target Export.xlsx - Excel',
                        'How_to_Manage.pptx  -  Protected View - PowerPoint',
                        'Sample presentation.pptx - PowerPoint',
                        'Sample presentation 2.pptx - PowerPoint']

    def __del__(self):
        self._savebinders()

    def _getfilename(self, win, app):
        # Remove the app name from the title
        appnamelen = len(app) + 3
        filename = win[:-appnamelen]

        # Split out the modifier (e.g., "AutoRecovered")
        matches = re.findall('(.*[\.][a-z]+)(?:  -  ([A-z\-0-9 ]+))*', filename)
        if matches is None:
            return matches
        return matches[0]

    def _winEnumHandler(self, hwnd, ctx ):
        if win32gui.IsWindowVisible( hwnd ):
            window = win32gui.GetWindowText( hwnd )
            self._winlist.append(window)

    def _getdoclist(self):
        if(self._mock):
            self._winlist = self._winlistmock
        else:
            self._winlist = []
            win32gui.EnumWindows( self._winEnumHandler, None )
    
        docdict = {}
        for win in self._winlist:
            for app in self._ahf.applist():
                if ' - ' + app in win:
                    matches = self._getfilename(win, app)
                    if not matches is None:
                        if len(matches) > 1:
                            modifier = matches[1]
                        else:
                            modifier = ""
                        
                        if app in self._ahf.applist():
                            path = self._ahf.gethandler(app).getpath(matches[0])
                        else:
                            path = ""
                          
                        docdict[matches[0]] = {"app": app, "winname": win, "path": path, "filename": matches[0], "modifier": modifier}
        
        return docdict

    def _workspacevalid(self, workspace):
        if self._workspaces.get(workspace) is None:
            print("No workspace {}".format(workspace))
            return False
        return True

    def _printworkspace(self, workspace):
        wsfiles = self._workspaces[workspace]
        print("Workspace {}".format(workspace))
        for idx, file in enumerate(wsfiles):
            print('{}: {}'.format(idx, file['filename']))

    def _savebinders(self):
        with open('docbinder.json','w') as output: 
            output.write(pprint.pformat(self._workspaces))

    def _openbinders(self):
        try:
            with open('docbinder.json') as json_file:
                self._workspaces = json.load(json_file)        
        except:
            # If JSON file doesn't exist
            pass

    def listdocs(self):
        self._doclist = self._getdoclist()
        self._docindex = list(self._doclist.keys())
        for idx, key in enumerate(self._doclist):
            doc = self._doclist[key]
            print("{}: {}    {}   [{}]".format(idx, doc['filename'], doc['modifier'], doc.get('workspace')))

    def clean(self, workspace):
        ''' Update the indicated workspace to reflect files that were closed '''
        if self._workspacevalid(workspace):
            print("Cleaning workspace {}".format(workspace))
            wsfiles = self._workspaces[workspace]
            docdict = self._getdoclist()
            for idx, wsfile in enumerate(wsfiles):
                if docdict.get(wsfile['filename']) is None:
                    print(" [REMOVING] {}".format(wsfile['filename']))
                    del wsfiles[idx]           

    def cleanall(self):
        ''' Cleans all active workspaces '''
        for ws in self._workspaces.keys():
            self.clean(ws)

    def update(self, workspace):
        ''' Update the indicated workspace by prompting to add files not in a list to a workspace '''
        pass
        
    def open(self, workspace):
        ''' Open the files belonging to the associated workspace '''
        if self._workspacevalid(workspace):
            print("Opening workspace {}".format(workspace))
            wsfiles = self._workspaces[workspace]
            for wsfile in wsfiles:
                if wsfile["app"] in self._ahf.applist():
                    self._ahf.gethandler(wsfile["app"]).openfile(wsfile["path"], wsfile["filename"])
         
    def openall(self):
        ''' Opens all saved workspaces '''
        for ws in self._workspaces.keys():
            self.open(ws)
        
    def add(self, workspace, filelist):
        ''' Add files to a workspace '''
        ''' db.add('CFO',(1,2,3)) '''
        if self._workspaces.get(workspace) is None:
            print("Created workspace {}".format(workspace))
            wsfiles = []
        else:
            print("Updating workspace {}".format(workspace))
            wsfiles = self._workspaces[workspace]

        for filenum in filelist:
            key = self._docindex[filenum]
            file = self._doclist[key]
            if file in wsfiles:
                print(' [SKIPPED] {} (already exists)'.format(file['filename']))
            else:
                wsfiles.append(file)
                print(' [ADDED] {}'.format(file['filename']))
        self._workspaces[workspace] = wsfiles

    def list(self, workspace = None):
        ''' List each workspace and its contents '''
        if workspace is None:
            if len(self._workspaces) > 0:
                for workspace in self._workspaces:
                    self._printworkspace(workspace)
            else:
                print("No workspaces created")
        elif self._workspacevalid(workspace):
            self._printworkspace(workspace)
        
    def delete(self, workspace):
        ''' Delete a workspace '''
        if self._workspacevalid(workspace):
            del self._workspaces[workspace]
            print('Deleted workspace {}'.format(workspace))

def _dbtest():
    print("Running _dbtest")
    db = DocBinder(mock=True)
    db.list()
    db.listdocs()
    db.add('CFO',(1,2,5))
    db.add('revenue',[1,2])
    db.list()
    db.delete('CFO')
    db.add('CFO',(1,2,5))
    db.add('CFO',(3,4))
    db.list('CFO')
    db.delete('CFO')

def _dbpersisttest():
    print("Running _dbpersisttest()")
    db = DocBinder(mock=True)
    db.listdocs()
    db.add('CFO',(1,2,5))
    db.add('revenue',[1,2])
    del db

    db2 = DocBinder(mock=True)
    db2.listdocs()
    db2.openall()
    db2.add('three',[3])

def _dbcleantest():
    db = DocBinder(mock=True)
    db.listdocs()
    db.add('CFO',(1,2,5))
    db.add('revenue',[1,2])
    del db._winlistmock[1]
    db.list()
    db.cleanall()
    db.list()
    db.clean('CFO')


if __name__ == "__main__":
    _dbcleantest()
