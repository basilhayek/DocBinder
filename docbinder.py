#TODO: Handle unsaved files
# DocBinder imports
import os
if os.name == "nt":
    import win32gui
    import win32com.client

import re

from apphandler import AppHandlerFactory

class DocBinder():
    
    _doclist = []
    _winlist = []
    _winlistmock = []
    _mock = False
    _workspaces = {}

    def __init__(self, mock = False):
        self._mock = mock
        self._ahf = AppHandlerFactory(mock)
        if(self._mock):
            self._winlistmock = ['210426%20Clilent%20Dashboard%20Mock-up%20draft (version 1).xlsb  -  AutoRecovered - Excel',
                        'Weekly Executive Report - Consolidated Numbers.xlsx - Excel',
                        '2021 Target Export.xlsx - Excel',
                        'How_to_Manage.pptx  -  Protected View - PowerPoint',
                        'Sample presentation.pptx  -  2 - PowerPoint',
                        'Sample presentation.pptx  -  1 - PowerPoint']

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
    
        
    
        doclist = []
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
                          
                        doclist.append({"app": app, "winname": win, "path": path, "filename": matches[0], "modifier": modifier})
        
        return doclist

    def _printworkspace(self, workspace):
        wsfiles = self._workspaces[workspace]
        print("Workspace {}".format(workspace))
        for idx, file in enumerate(wsfiles):
            print('{}: {}'.format(idx, file['filename']))

    def listdocs(self):
        self._doclist = self._getdoclist()
        for idx, doc in enumerate(self._doclist):
            print("{}: {}    {}   [{}]".format(idx, doc['filename'], doc['modifier'], doc.get('workspace')))

    def clean(self, workspace = None):
        ''' Update the indicated workspace to reflect files that were closed '''
        ''' Passing in an empty string cleans the last workspace '''
        ''' { "workspaces": [{"workspace":"cfo","files":[{"app":"Excel","filename":"filename","path":"path"}]}]} '''
        pass

    def update(self, workspace = None):
        ''' Update the indicated workspace by prompting to add files not in a list to a workspace '''
        ''' Passing in an empty string updates the last workspace '''
        pass
        
    def open(self, workspace):
        ''' Open the files belonging to the associated workspace '''
        if self._workspaces.get(workspace) is None:
            print("No workspace {}".format(workspace))
            return

        print("Opening workspace {}".format(workspace))
        wsfiles = self._workspaces[workspace]
        for wsfile in wsfiles:
            if wsfile["app"] in self._ahf.applist():
                self._ahf.gethandler(wsfile["app"]).openfile(wsfile["path"], wsfile["filename"])
         
    def openall(self):
        ''' Opens all saved workspaces '''
        pass
        
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
            file = self._doclist[filenum]
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
        else:
            if workspace in self._workspaces:
                self._printworkspace(workspace)
            else:
                print('Workspace {} does not exist'.format(workspace))
        
    def delete(self, workspace):
        ''' Delete a workspace '''
        if workspace in self._workspaces:
            del self._workspaces[workspace]
            print('Deleted workspace {}'.format(workspace))

def _dbtest():
    print("Running tests")
    db = DocBinder(mock=True)
    # print(db._getdoclist())
    db.list()
    db.listdocs()
    db.add('CFO',(1,2,5))
    db.list()
    db.delete('CFO')
    db.add('CFO',(1,2,5))
    db.add('CFO',(3,4))
    db.list('CFO')
    db.delete('CFO')

if __name__ == "__main__":
    _dbtest()
