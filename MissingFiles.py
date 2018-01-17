# Copyright (C) 2002 - 2017 SpurgeonumBars Woods LLC
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of version 2 of the GNU General Public License as
# published by the Free Software Foundation.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
#

"""This module identifies missing external files and helps fix the problem. """

__author__ = 'David K. Woods <dwoods@transana.com>'

# Import wxPython
import wx

# Import Transana's Dialogs
import Dialogs
# Import Transana's Database Interface
import DBInterface
# Import File Management Tool
import FileManagement
# Import Transana's Local File Transfer module
import LocalFileTransfer
# Import Transana's Constants
import TransanaConstants
# Import Transana's Globals
import TransanaGlobal

# import Python's os and sys modules
import os, sys
# import Python's shutil module
import shutil

class AutoWidthListCtrl(wx.ListCtrl, wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin):
    """ Create an AutoWidth List Control using mixins """
    def __init__(self, parent, ID, style=0):
        """ Initialize the AutoWidth List Control """
        # Initialize the ListCtrl
        wx.ListCtrl.__init__(self, parent, ID, style=style)
        # Add the AutoWidth mixin
        wx.lib.mixins.listctrl.ListCtrlAutoWidthMixin.__init__(self)

class MissingFiles(wx.Dialog, wx.lib.mixins.listctrl.ColumnSorterMixin):
    """ Create a tool for handling missing files"""
    def __init__(self, parent, title, controlObject):
        # Note the Control Object passed in, used in the Help call
        self.ControlObject = controlObject
        # Get the Media Library Directory
        mediaDir = TransanaGlobal.configData.videoPath.replace('\\', '/')
        # We need a specifically-named dictionary in a particular format for the ColumnSorterMixin.
        # Initialize that here.  (This also tells us if we've gotten the data yet!)
        self.itemDataMap = {}

        # Create the Form
        wx.Dialog.__init__(self,parent,title=title,size=(1000,650), style=wx.CAPTION | wx.RESIZE_BORDER | wx.CLOSE_BOX)
        
        # Create a Sizer
        s1 = wx.BoxSizer(wx.VERTICAL)
        # Create instructions
        prompt  = _("The following files are not where Transana expects them to be.  First, please select one or more files and search for them using the Search button below.\n")
        prompt += _("If a copy of the file is found, You can instruct Transana to use it where it is, to copy it to where Transana expects it, or to move it.  If no copy of the file \n")
        prompt += _("is found, you can change the Search Directory and search again.")
        # Place the instructions on the form
        txt = wx.StaticText(self, -1, prompt)
        # Add the instructions to the Sizer
        s1.Add(txt, 0, wx.EXPAND | wx.ALL, 10)
        
        # Create the list of missing files
        self.fileList = AutoWidthListCtrl(self, -1, style=wx.LC_REPORT | wx.BORDER_SUNKEN | wx.LC_SORT_ASCENDING)
        # Add in the Column Sorter mixin to make the results panel sortable
        wx.lib.mixins.listctrl.ColumnSorterMixin.__init__(self, 3)
        # Add the missing file list to the Sizer
        s1.Add(self.fileList, 1, wx.EXPAND)

        # Create a horizontal sizer
        h1 = wx.BoxSizer(wx.HORIZONTAL)
        # Create a Search button and give it a handler
        self.btnSearch = wx.Button(self, -1, _("Search for File(s)"))
        self.btnSearch.Bind(wx.EVT_BUTTON, self.OnSearch)
        # Add it to the Button Sizer
        h1.Add(self.btnSearch, 0, wx.TOP | wx.LEFT | wx.RIGHT, 10)
        # Add a Label for the Search Directory
        searchTxt = wx.StaticText(self, -1, _("Starting Search Directory"))
        # Add the label to the Sizer
        h1.Add(searchTxt, 0, wx.ALIGN_CENTER_VERTICAL | wx.TOP | wx.LEFT | wx.RIGHT, 10)
        # Add a TextCtrl for the Search Path
        self.path = wx.TextCtrl(self, -1, mediaDir)
        # Add the Search Path to the Sizer
        h1.Add(self.path, 1, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 10)
        # Create a Browse Button
        browse = wx.Button(self, -1, _("Browse"))
        # Add the Browse button to the Sizer
        h1.Add(browse, 0, wx.TOP | wx.LEFT | wx.RIGHT, 10)
        # Add the Button Press Handler
        browse.Bind(wx.EVT_BUTTON, self.OnBrowse)
        # Add the horizontal Sizer to the Main Sizer
        s1.Add(h1, 0, wx.EXPAND)
        
        # Create a Sizer for the Button Bar
        h2 = wx.BoxSizer(wx.HORIZONTAL)
        # Create a Use Existing Location button and give it a handler
        self.btnUse = wx.Button(self, -1, _("Use Actual Location"))
        self.btnUse.Bind(wx.EVT_BUTTON, self.OnUseCopyMove)
        # Add it to the Button Sizer
        h2.Add(self.btnUse, 0, wx.ALL, 10)
        # Create a Copy File button and give it a handler
        self.btnCopy = wx.Button(self, -1, _("Copy File"))
        self.btnCopy.Bind(wx.EVT_BUTTON, self.OnUseCopyMove)
        # Add it to the Button Sizer
        h2.Add(self.btnCopy, 0, wx.ALL, 10)
        # Create a Move File button and give it a handler
        self.btnMove = wx.Button(self, -1, _("Move File"))
        self.btnMove.Bind(wx.EVT_BUTTON, self.OnUseCopyMove)
        # Add it to the Button Sizer
        h2.Add(self.btnMove, 0, wx.ALL, 10)
        # Create a Rename button and give it a handler
        self.btnRename = wx.Button(self, -1, _("Change File Name"))
        self.btnRename.Bind(wx.EVT_BUTTON, self.OnUseCopyMove)
        # Add it to the Button Sizer
        h2.Add(self.btnRename, 0, wx.ALL, 10)
        # Create a File Management button and give it a handler
        self.btnFileMgt = wx.Button(self, -1, _("File Management"))
        self.btnFileMgt.Bind(wx.EVT_BUTTON, self.OnFileManagement)
        # Add it to the Button Sizer
        h2.Add(self.btnFileMgt, 0, wx.ALL, 10)
        # Add an expandable spacer to the Button Sizer
        h2.Add((1, 1), 1, wx.EXPAND)
        # Create an OK button and give it a handler
        self.btnOK = wx.Button(self, -1, _("OK"))
        self.btnOK.Bind(wx.EVT_BUTTON, self.OnOk)
        # Add it to the Button Sizer
        h2.Add(self.btnOK, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        # Create a Help button and give it a handler
        self.btnHelp = wx.Button(self, -1, _("Help"))
        self.btnHelp.Bind(wx.EVT_BUTTON, self.OnHelp)
        # Add it to the Button Sizer
        h2.Add(self.btnHelp, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
        # Add the Button Sizer to the Main Sizer
        s1.Add(h2, 0, wx.EXPAND)
        
        # Set the main sizer
        self.SetSizer(s1)
        # Lay out the Form
        self.Layout()
        self.SetAutoLayout(True)
        # Center the form on screen
        TransanaGlobal.CenterOnPrimary(self)
        # Populate the Missing Files List
        self.UpdateMissingFileList()

    def UpdateMissingFileList(self):
        """ (Re)populate the Missing Files List """
        # Clear the Missing File List
        self.fileList.ClearAll()
        # Define columns for missing files, their expected location and the location where they are actually found
        self.fileList.InsertColumn(0, _("Missing File"))
        self.fileList.InsertColumn(1, _("Expected Location"))
        self.fileList.InsertColumn(2, _("Actual Location"))

        # Initialize the item data dictionary used for the sorter
        self.itemDataMap = {}
        # Initialize a counter which serves as the key to the item data dictionary
        counter = 0
        # Get a list of all missing external files (Media and Image files) from the Database
        fileList = DBInterface.GetMissingFilesList()
        # Iterate through the list of missing files
        for filename in fileList:
            # ... separate the Path and the File Name
            (fDir, fName) = os.path.split(filename)
            # Create a new Row in the File List, placing the File Name in the first Column
            index = self.fileList.InsertStringItem(sys.maxint, fName)
            # Add the Expected Path (where the file is NOT!) in the second column.  Leave the third column blank for now.
            self.fileList.SetStringItem(index, 1, fDir)
            # Add the data to the item data dictionary for the sorter
            self.itemDataMap[counter] = (fName, fDir, '')
            # Set the Missing File List's item data
            self.fileList.SetItemData(index, counter)
            # Increment the counter
            counter += 1
        # Sort by the first column
        self.SortListItems(0)
        # Column widths have proven more complicated than I'd like due to cross-platform differences.
        # This formula seems to work pretty well.
        # If there are no Files included ...
        if self.fileList.GetItemCount() == 0:
            # ... use the Headers to set Column Widths
            self.fileList.SetColumnWidth(0, wx.LIST_AUTOSIZE_USEHEADER)
            self.fileList.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        # If there are files listed ...
        else:
            # ... use the data.  (On MacOS, wx.LIST_AUTOSIZE_USEHEADER ignores larger data below the header!)
            self.fileList.SetColumnWidth(0, wx.LIST_AUTOSIZE)
            self.fileList.SetColumnWidth(1, wx.LIST_AUTOSIZE)

    def find_all(self, name, path):
        """ Given a File Name and a Starting Path, find all instances of that file name within the path and subdirectories """
        # Initialize the list of results
        result = []
        # Get the lower case version of the name to avoid case sensitivity
        namelo = name.lower()
        # Search all directories from the starting path onwards
        for root, dirs, files in os.walk(path):
            # For each file listed in the directory ...
            for x in range(len(files)):
                # ... convert the file name to its lower case version
                files[x] = files[x].lower()
            # If the file being sought is in this directory's list of files ...
            if namelo in files:
                # ... add the full file path to the list of results
                result.append(os.path.join(root, name).replace('\\', '/'))
        # Return the list of matches found
        return result

    def OnBrowse(self, event):
        prompt = _("Select a directory")
        # Create a Directory Dialog Box
        dlg = wx.DirDialog(self, prompt, self.path.GetValue(), style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        # Display the Dialog modally and process if OK is selected.
        if dlg.ShowModal() == wx.ID_OK:
            self.path.SetValue(dlg.GetPath())
        # Destroy the Dialog
        dlg.Destroy

    def OnSearch(self, event):
        """ Event Handler for the Search Button """
        # Get the first selected element in the Missing Files list
        sel = self.fileList.GetFirstSelected()
        # If not items are selected in the Missing Files list, but there IS data in the list ...
        if (sel == -1) and (self.fileList.GetItemCount() > 0):
            # ... select the first item in the list ...
            self.fileList.Select(0)
            # ... and indicate that the selection has been made
            sel = 0
        # While there are still selected items ...
        while sel > -1:
            # ... search for the selection's file name in the search path
            results = self.find_all(self.fileList.GetItem(sel, 0).GetText(), self.path.GetValue())
            # If NO results are found ...
            if len(results) == 0:
                # ... indicate that the file was not found by defining the message to be placed in the Missing Files list
                path = _("File Not Found")
            # If ONE result was found ...
            elif len(results) == 1:
                # ... then we know what data to add (path only, excluding file name) to the Missing Files list ...
                path, fn = os.path.split(results[0])
            # If more than one copy of the file is found in the search path ...
            else:
                # ... create a Temporary Dialog showing all files found and asking the user to make a single selection
                tmpDlg = wx.SingleChoiceDialog(self, _("Choose which file to use"), _("Multiple Files Found"), results, wx.CHOICEDLG_STYLE)
                # Show this temporary Dialog.  If OK is pressed ...
                if tmpDlg.ShowModal() == wx.ID_OK:
                    # ... then we know what data to add (path only, excluding file name) to the Missing Files list ...
                    path, fn = os.path.split(tmpDlg.GetStringSelection())
                # If the dialog is cancelled ...
                else:
                    # ... we have nothing to put into the Missing Files list
                    path = ''
                # Destroy the Temporary Dialog
                tmpDlg.Destroy()
            # Add the resulting information to the Missing Files list in the Actual Path column
            self.fileList.SetStringItem(sel, 2, path)

            # Get the item's Sort Key
            counter = self.fileList.GetItemData(sel)
            # Update the item data dictionary for the sorter
            self.itemDataMap[counter] = (self.itemDataMap[counter][0], self.itemDataMap[counter][1], path)

            
            # Make sure the list is properly updated on screen
            self.fileList.Update()
            # Move on to the next selection in the Missing Files List, if there is one
            sel = self.fileList.GetNextSelected(sel)

    def OnUseCopyMove(self, event):
        """ Handle "Use Actual Location", "Copy File" and "Move File" Button Presses """
        # Create an empty list of the entries which should be deleted from the Missing Files list when we are finished here
        entriesToDelete = []

        # Get the first selected element in the Missing Files list
        sel = self.fileList.GetFirstSelected()
        # If there are no selections in the list ...
        if sel == -1:
            # ... there's nothing to do here.
            return

        # While there are still selections in the list ...
        while sel > -1:
            # If the Actual Location is not defined for a file ...
            if self.fileList.GetItem(sel, 2).GetText() == '':
                # ... call the Search for that file name automatically!
                self.OnSearch(event)

            # Get the File Name, the Existing Path, Actual Path and the for the current selection from the Missing File List
            fileName = self.fileList.GetItem(sel, 0).GetText()
            existingPath = self.fileList.GetItem(sel, 1).GetText()
            actualPath = self.fileList.GetItem(sel, 2).GetText()

            # If we're on Windows ...
            if 'wxMSW' in wx.PlatformInfo:
                # ... replace the File Separator
                fileName = fileName.replace('/', os.sep)
                existingPath = existingPath.replace('/', os.sep)
                actualPath = actualPath.replace('/', os.sep)

            # If the file was not found ...
            if (actualPath == _("File Not Found")) and (event.GetId() != self.btnRename.GetId()):
                # ... we can't do anything here.
                pass
            # If the "Use Actual Location" button was pressed ...
            elif event.GetId() == self.btnUse.GetId():
                # ... update all copies of the file in the Database with the new File Path
                if not DBInterface.UpdateDBFilenames(self, actualPath, [fileName]):
                    # If some or all of the records could not be updated, display an error message
                    infodlg = Dialogs.InfoDialog(self, _('Update Failed.  Some records that would be affected may be locked by another user.'))
                    infodlg.ShowModal()
                    infodlg.Destroy()
                # If the update succeeeded ...
                else:
                    # ... we can remove this item from the Missing Files list
                    entriesToDelete.append(sel)
            # If the "Copy File" button was pressed ...
            elif event.GetId() == self.btnCopy.GetId():
                # ... default to signalling success
                success = True
                # Combine the actual path and the file name
                sourceFile = os.path.join(actualPath, fileName)
                # If the path that should receive the copy doesn't exist ...
                if not os.path.exists(existingPath):
                    # create it
                    os.makedirs(existingPath)
                # Check if the destination file already exists.  If NOT ...
                if not os.path.exists(os.path.join(existingPath, fileName)):
                    # If the file is less than 10 MB ...
                    if os.stat(sourceFile)[6] < 10000000:
                        # ... start exception handling
                        try:
                            # Copy the file to the destination path using the fast shutil module
                            shutil.copyfile(sourceFile, os.path.join(existingPath, fileName))
                        # If an exception is raised ...
                        except:
                            # ... the copy failed.
                            success = False

                            print sys.exc_info()[0]
                            print sys.exc_info()[1]
                    
                    # For larger files ...
                    else:
                        # ... let's copy using Transana's LocalFileTransfer module, which gives a progress bar
                        dlg = LocalFileTransfer.LocalFileTransfer(self, _("Copy File"), sourceFile, existingPath)
                        success = dlg.TransferSuccessful()
                        dlg.Destroy()
                # If the file DOES exist ...  (The file should not even be listed if this is true!!!!)
                else:
                    # ... then the copy failed
                    success = False

                # if the copy succeeded ...
                if success:
                    # ... we can remove the file name from the Missing Files list
                    entriesToDelete.append(sel)
                # If the copy failed ...
                else:
                    # ... display an error message
                    errMsg = _('File "%s" could not be copied.')
                    errData = (sourceFile,)
                    if os.path.exists(os.path.join(existingPath, fileName)):
                        errMsg += '\n' + _('File "%s" already exists.')
                        errData += (os.path.join(existingPath, fileName),)
                    if 'unicode' in wx.PlatformInfo:
                        errMsg = unicode(errMsg, 'utf8')
                    # Display the error message
                    dlg = Dialogs.ErrorDialog(self, errMsg % errData)
                    dlg.Show()  # NOT Modal, and no Destroy.  This allows user interruption of the copy.
            # If the Move File button is pressed ...
            elif event.GetId() == self.btnMove.GetId():
                # Start Exception Handling
                try:
                    # Check if the destination file already exists.  If NOT ...
                    if not os.path.exists(os.path.join(existingPath, fileName)):
                        # ... os.renames accomplishes a fast LOCAL Move
                        os.renames(os.path.join(actualPath, fileName), os.path.join(existingPath, fileName))
                        # If that worked, signal that we should remove this file from the Missing Files list
                        entriesToDelete.append(sel)
                    # If the file DOES exist ...  (The file should not even be listed if this is true!!!!)
                    else:
                        # Create an error message
                        errMsg = _('File "%s" could not be moved.') + '\n' + _('File "%s" already exists.')
                        if 'unicode' in wx.PlatformInfo:
                            errMsg = unicode(errMsg, 'utf8')
                        # Display the error message
                        dlg = Dialogs.ErrorDialog(self, errMsg % (os.path.join(actualPath, fileName), os.path.join(existingPath, fileName)))
                        dlg.Show()  # NOT Modal, and no Destroy prevents interruption
                # Process exceptions
                except:

                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    
                    # Create an error message
                    errMsg = _('File "%s" could not be moved.')
                    if 'unicode' in wx.PlatformInfo:
                        errMsg = unicode(errMsg, 'utf8')
                    # Display the error message
                    dlg = Dialogs.ErrorDialog(self, errMsg % fileName)
                    dlg.Show()  # NOT Modal, and no Destroy.
            # If the Change File Name button is pressed ...
            elif event.GetId() == self.btnRename.GetId():
                # Determine the File's extension
                (fn, ext) = os.path.splitext(fileName)
                # If we have a known File Type or if blank, use "All Media Files".
                # If it's an unrecognized type, go to "All Files"
                if (TransanaGlobal.configData.LayoutDirection == wx.Layout_LeftToRight) and \
                   (ext.lower() in ['.mpg', '.avi', '.mov', '.mp4', '.m4v', '.wmv', '.mp3', '.wav', '.wma', '.aac', '']):
                    fileType =  '*.mpg;*.avi;*.mov;*.mp4;*.m4v;*.wmv;*.mp3;*.wav;*.wma;*.aac'
                elif (TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft) and \
                   (ext.lower() in ['.mpg', '.avi', '.wmv', '.mp3', '.wav', '.wma', '.aac', '']):
                    fileType =  '*.mpg;*.avi;*.wmv;*.mp3;*.wav;*.wma;*.aac'
                else:
                    fileType = ''
                # Invoke the File Selector with the proper default directory, filename, file type, and style
                fs = wx.FileSelector(_("Select the media file"),
                                existingPath,
                                fileName,
                                fileType, 
                                _(TransanaConstants.fileTypesString), 
                                wx.OPEN | wx.FILE_MUST_EXIST)
                # If user didn't cancel ..
                if fs != "":
                    # Mac Filenames use a different encoding system.  We need to adjust the string returned by the FileSelector.
                    # Surely there's an easier way, but I can't figure it out.
                    if 'wxMac' in wx.PlatformInfo:
                        fs = Misc.convertMacFilename(fs)
                    # Split the new file name into path and file name
                    (actualPath, newFileName) = os.path.split(fs)
                    # ... update all copies of the file in the Database with the new File Path
                    if not DBInterface.UpdateDBFilenames(self, actualPath, [fileName], newName=newFileName):
                        # If some or all of the records could not be updated, display an error message
                        infodlg = Dialogs.InfoDialog(self, _('Update Failed.  Some records that would be affected may be locked by another user.'))
                        infodlg.ShowModal()
                        infodlg.Destroy()
                    # If the update succeeeded ...
                    else:
                        # ... we can remove this item from the Missing Files list
                        entriesToDelete.append(sel)

            # Move on to the next selection in the Missing Files List, if there is one
            sel = self.fileList.GetNextSelected(sel)

        # Sort the list of items to be deleted in Reverse order so the deletes don't change the item numbers as we go.
        entriesToDelete.sort(reverse=True)
        # For each item to be deleted ...
        for entry in entriesToDelete:
            # ... delete the item from the Missing Files list.
            self.fileList.DeleteItem(entry)
        # Update the control on screen
        self.fileList.Update()

    def OnFileManagement(self, event):
        # Create a File Management Window
        self.fileManagementWindow = FileManagement.FileManagement(self, -1, _("Transana File Management"))
        # Set up, display, and process the File Management Window
        self.fileManagementWindow.Setup(showModal=True)
        # File information may have changed.  Update the Missing File list on screen.
        self.UpdateMissingFileList()

    def OnOk(self, event):
        """ Handle the OK button press """
        # Just close the form.
        self.Close()

    def OnHelp(self, event):
        """ Handle the Help Button press """
        # ... call Help!
        self.ControlObject.Help("Find Missing Files")

    def GetListCtrl(self):
        """ Pointer to the Missing Files List control, required for the ColumnSorterMixin """
        # Return a pointer to the Missing Files List control
        return self.fileList
