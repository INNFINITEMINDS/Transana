# Copyright (C) 2002 - 2017 Spurgeon Woods LLC
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

""" This module implements the Spreadsheet Data Import function for Transana """

__author__ = 'David Woods <dwoods@transana.com>'

DEBUG = False
if DEBUG:
    print "SpreadsheetDataImport DEBUG is ON!!"

# Import wxPython
import wx
# Import wxPython's Wizard
import wx.wizard as wiz

# Import Transana's Collection object
import Collection
# Import Transana's DBInterface module
import DBInterface
# Import Transana's Dialogs
import Dialogs
# Import Transana's Document Object
import Document
# Import Transana's Keyword Object
import KeywordObject
# Import Transana's Library Object
import Library
# Import Transana's Quote object
import Quote
# Import Transana's Constants
import TransanaConstants
# Import Transana's Exceptions
import TransanaExceptions
# Import Transana's Global Variables
import TransanaGlobal

# import Python's csv module, which reads Comma Separated Values files
import csv
# import Python's datetime module
import datetime
# Import Python's os module
import os
# Import Python's sys module
import sys


class EditBoxFileDropTarget(wx.FileDropTarget):
    """ This simple derived class let's the user drop files onto an edit box """
    def __init__(self, editbox):
        """ Initialize a File Drop Target """
        # Initialize the FileDropTarget object
        wx.FileDropTarget.__init__(self)
        # Make the Edit Box passed in a File Drop Target
        self.editbox = editbox
        
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        # Insert the FIRST file name into the edit box
        self.editbox.SetValue(files[0])


class WizPage(wiz.PyWizardPage):
    """ Base class for individual wizard pages.  Provides:
          Title
          Back / Forward / Cancel button """
    def __init__(self, parent, title):
        """ Initialize the Wizard Page """
        # Initialize the Previous and Next pointer to None
        self.prev = self.next = None
        # Remember the parent
        self.parent = parent

        # Initialize the PyWizardPage object
        wiz.PyWizardPage.__init__(self, parent)
        # Define a main Sizer for the page
        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # Display the Page Title
        title = wx.StaticText(self, -1, title)
        title.SetFont(wx.Font(14, wx.SWISS, wx.NORMAL, wx.BOLD))
        self.sizer.Add(title, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        self.sizer.Add(wx.StaticLine(self, -1), 0, wx.EXPAND|wx.ALL, 5)

        # Set the main Sizer
        self.SetSizer(self.sizer)

        # Identify the Previous and Next buttons so they can be manipulated
        self.prevButton = parent.FindWindowById(wx.ID_BACKWARD)
        self.nextButton = parent.FindWindowById(wx.ID_FORWARD)

    def GetPrev(self):
        """ Get Previous Page function """
        return self.prev

    def SetPrev(self, prev):
        """ Set Previous Page function """
        self.prev = prev

    def GetNext(self):
        """ Get Next Page function """
        return self.next

    def SetNext(self, next):
        """ Set Next Page function """
        self.next = next

    def IsComplete(self):
        """ Has this page been completed?  Defaults to False, must be over-ridden! """
        return False


class GetFileNamePage(WizPage):
    """ Get File Name wizard page """
    def __init__(self, parent, title):
        """ Define the Wizard Page that gets the File to be imported """
        # Inherit from WizPage
        WizPage.__init__(self, parent, title)

        # Add the Source File label
        lblSource = wx.StaticText(self, -1, _("Source Data File:"))
        self.sizer.Add(lblSource, 0, wx.TOP | wx.LEFT | wx.RIGHT, 10)

        # Create the box1 sizer, which will hold the source file and its browse button
        box1 = wx.BoxSizer(wx.HORIZONTAL)

        # Create the Source File text box
        self.txtSrcFileName = wx.TextCtrl(self, -1)
        # Make the Source File a File Drop Target
        self.txtSrcFileName.SetDropTarget(EditBoxFileDropTarget(self.txtSrcFileName))

        # Handle ALL changes to the source filename
        self.txtSrcFileName.Bind(wx.EVT_TEXT, self.OnSrcFileNameChange)
        # Add the text box to the sizer
        box1.Add(self.txtSrcFileName, 1, wx.EXPAND)
        # Spacer
        box1.Add((4, 0))
        # Create the Source File Browse button
        self.srcBrowse = wx.Button(self, -1, _("Browse"))
        self.srcBrowse.Bind(wx.EVT_BUTTON, self.OnBrowse)
        box1.Add(self.srcBrowse, 0, wx.LEFT, 10)
        # Add the Source Sizer to the Main Sizer
        self.sizer.Add(box1, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 10)

        # Add Encoding label
        lblEncoding = wx.StaticText(self, -1, _("File Encoding:"))
        self.sizer.Add(lblEncoding, 0, wx.TOP | wx.LEFT | wx.RIGHT, 10)

        # Add Encoding selection
        choices = ["ascii", "big5", "big5hkscs", "cp037", "cp424", "cp437", "cp500", "cp720", "cp737", "cp775", "cp850", "cp852", "cp855", "cp856", "cp857",
                   "cp858", "cp860", "cp861", "cp862", "cp863", "cp864", "cp865", "cp866", "cp869", "cp874", "cp875", "cp932", "cp949", "cp950", "cp1006",
                   "cp1026", "cp1140", "cp1250", "cp1251", "cp1252", "cp1253", "cp1254", "cp1255", "cp1256", "cp1257", "cp1258", "euc_jp", "euc_jis_2004",
                   "euc_jisx0213", "euc_kr", "gb2312", "gbk", "gb18030", "hz", "iso2022_jp", "iso2022_jp_1", "iso2022_jp_2", "iso2022_jp_2004", "iso2022_jp_3",
                   "iso2022_jp_ext", "iso2022_kr", "latin_1", "iso8859_2", "iso8859_3", "iso8859_4", "iso8859_5", "iso8859_6", "iso8859_7", "iso8859_8",
                   "iso8859_9", "iso8859_10", "iso8859_13", "iso8859_14", "iso8859_15", "iso8859_16", "johab", "koi8_r", "koi8_u", "mac_cyrillic",
                   "mac_greek", "mac_iceland", "mac_latin2", "mac_roman", "mac_turkish", "ptcp154", "shift_jis", "shift_jis_2004", "shift_jisx0213",
                   "utf_32", "utf_32_be", "utf_32_le", "utf_16", "utf_16_be", "utf_16_le", "utf_7", "utf_8", "utf_8_sig"]
        self.txtSrcEncoding = wx.Choice(self, -1, choices = choices)
        self.txtSrcEncoding.SetStringSelection('cp1252')
        self.sizer.Add(self.txtSrcEncoding, 0, wx.EXPAND | wx.TOP | wx.LEFT | wx.RIGHT, 10)

        self.sizer.Add((1, 24))
        prompt = _("For more information on file formats and encoding, please see the Help page.")
        info = wx.StaticText(self, -1, prompt)
        font = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        info.SetFont(font)
        self.sizer.Add(info, 0, wx.TOP | wx.LEFT | wx.RIGHT, 10)

    def IsComplete(self):
        """ IsComplete signals whether an EXISTING file has been selected """
        return os.path.exists(self.txtSrcFileName.GetValue())

    def OnBrowse(self, event):
        """ Browse Button event handler """
        # Get Transana's File Filter definitions
        fileTypesString = _("All supported files (*.csv, *.txt)|*.csv;*.txt|Comma Separated Values files (*.csv)|*.csv|Tab Delimited Text files (*.txt)|*.txt|All files (*.*)|*.*")
        # Create a File Open dialog.
        fs = wx.FileDialog(self, _('Select a spreadsheet data file:'),
                        TransanaGlobal.configData.videoPath,
                        "",
                        fileTypesString, 
                        wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        # Select "All supported files" as the initial Filter
        fs.SetFilterIndex(0)
        # Show the dialog and get user response.  If OK ...
        if fs.ShowModal() == wx.ID_OK:
            # ... get the selected file name
            self.fileName = fs.GetPath()
        # If not OK ...
        else:
            # ... signal Cancel by return a blank file name
            self.fileName = ''

        # Destroy the File Dialog
        fs.Destroy()

        # Add the selected File Name to the File Name text box
        self.txtSrcFileName.SetValue(self.fileName)

        # Determine the file extension
        (filename, extension) = os.path.splitext(self.fileName)
        # Set the default encoding, based on the file extension selected (In Excel, CSV is CP1252 encoded, Unicode TXT is UTF-16 encoded.)
        if extension.lower() == '.csv':
            self.txtSrcEncoding.SetStringSelection('cp1252')
        elif extension.lower() == '.txt':
            self.txtSrcEncoding.SetStringSelection('utf_16')

    def OnSrcFileNameChange(self, event):
        """ Process changes to the File Name Text Box """
        # If we have a valid, existing file ...
        if self.IsComplete():
            # ... enable the Next button
            self.nextButton.Enable(True)
        # If we don't have a valid file ...
        else:
            # ... disable the Next button
            self.nextButton.Enable(False)


class GetRowsOrColumnsPage(WizPage):
    """ Wizard page to find out if the data is organized by Rows or by Columns """
    def __init__(self, parent, title):
        """ Define the Wizard Page that gets the File's Data Orientation """
        # Inherit from WizPage
        WizPage.__init__(self, parent, title)

        # Add a Text Field for the File Name
        self.fileName = wx.StaticText(self, -1, '')
        self.sizer.Add(self.fileName, 0, wx.ALL, 5)
        self.sizer.Add((1, 5))

        # Add a Hozizontal Box Sizer
        boxH = wx.BoxSizer(wx.HORIZONTAL)
        
        # A Vertical Box Sizer 1 (left)
        boxV1 = wx.BoxSizer(wx.VERTICAL)
        # Add direction prompt
        prompt1 = wx.StaticText(self, -1, _('Contents of the first Row'))
        boxV1.Add(prompt1, 0)
        # Add a multi-line TextCtrl to hold first Row items
        self.txt1 = wx.TextCtrl(self, -1, style=wx.TE_MULTILINE)
        boxV1.Add(self.txt1, 1, wx.EXPAND | wx.ALL, 3)
        # Add a checkbox for selecting ROWS
        self.chkRows = wx.CheckBox(self, -1, " " + _("Prompts shown in first Row"), style=wx.CHK_2STATE)
        boxV1.Add(self.chkRows, 0)

        # A Vertical Box Sizer 2 (right)
        boxV2 = wx.BoxSizer(wx.VERTICAL)
        # Add direction prompt
        prompt2 = wx.StaticText(self, -1, _('Contents of the first Column'))
        boxV2.Add(prompt2, 0)
        # Add a multi-line TextCtrl to hold first Column items
        self.txt2 = wx.TextCtrl(self, -1, style=wx.TE_MULTILINE)
        boxV2.Add(self.txt2, 1, wx.EXPAND | wx.ALL, 3)
        # Add a checkbox for selecting COLUMNS
        self.chkColumns = wx.CheckBox(self, -1, " " + _("Prompts shown in first Column"), style=wx.CHK_2STATE)
        boxV2.Add(self.chkColumns, 0)

        # Add the Vertical Sizers to the Horizontal Sizer
        boxH.Add(boxV1, 1, wx.EXPAND)
        boxH.Add(boxV2, 1, wx.EXPAND)
        # Add the Horizontal Sizer to the Main Sizer
        self.sizer.Add(boxH, 1, wx.EXPAND)

        # Set the processor for CheckBoxes
        self.Bind(wx.EVT_CHECKBOX, self.OnCheckbox)

    def IsComplete(self):
        """ IsComplete signals whether either Checkbox has been checked """
        return self.chkColumns.GetValue() or self.chkRows.GetValue()

    def OnCheckbox(self, event):
        """ Process Checkbox Activity """
        # If the Columns Checkbox is clicked ...
        if event.GetId() == self.chkColumns.GetId():
            # ... and has been CHECKED ...
            if self.chkColumns.GetValue():
                # ... un-check the Rows checkbox ...
                self.chkRows.SetValue(False)
                # ... and enable the Next Button
                self.nextButton.Enable(True)
            # ... and has been UN-CHECKED ...
            else:
                # ... disable the Next Button
                self.nextButton.Enable(False)
        # If the Rows Checkbox is clicked ...
        else:
            # ... and has been CHECKED ...
            if self.chkRows.GetValue():
                # ... un-check the Columns checkbox ...
                self.chkColumns.SetValue(False)
                # ... and enable the Next Button
                self.nextButton.Enable(True)
            # ... and has been UN-CHECKED ...
            else:
                # ... disable the Next Button
                self.nextButton.Enable(False)


class GetItemsToIncludePage(WizPage):
    """ Wizard page to find out which items to include and how to present them """
    def __init__(self, parent, title):
        """ Define the Wizard Page that information about organizing and including data """
        # Inherit from WizPage
        WizPage.__init__(self, parent, title)

        # Add the first Instruction Text
        instructions1 = _("Please select the item used to uniquely identify Participants")
        txtInstructions1 = wx.StaticText(self, -1, instructions1)
        self.sizer.Add(txtInstructions1, 0, wx.ALL, 5)

        # Add a Choice Box for Unique Identifier
        self.identifier = wx.Choice(self, -1)
        self.sizer.Add(self.identifier, 0, wx.EXPAND | wx.ALL, 5)

        # Add a spacer
        self.sizer.Add((1, 5))

        # Add the second Instruction Text
        instructions2 = _("Please select which items to include in the Transana Documents to be created.")
        txtInstructions2 = wx.StaticText(self, -1, instructions2)
        self.sizer.Add(txtInstructions2, 0, wx.ALL, 5)

        # Add a ListBox for the Prompts / Questions with multi-select enabled
        self.questions = wx.ListBox(self, -1, style = wx.LB_EXTENDED)
        self.sizer.Add(self.questions, 1, wx.EXPAND | wx.ALL, 5)
        # Bind a handler for Item Selection
        self.Bind(wx.EVT_LISTBOX, self.OnItemSelect)

        # Add a spacer
        self.sizer.Add((1, 5))

    def IsComplete(self):
        """ IsComplete signals whether ANY questions have been selected """
        return len(self.questions.GetSelections()) > 0

    def OnItemSelect(self, event):
        """ Handler for Question / Prompt Selection """
        # Check Next Button Enabling on each change
        self.nextButton.Enable(self.IsComplete())

    
class GetAutoCodePage(WizPage):
    """ Wizard page to find out which items to auto-code """
    def __init__(self, parent, title):
        """ Define the Wizard Page that gets the auto-code information """
        # Inherit from WizPage
        WizPage.__init__(self, parent, title)

        # Add a spacer
        self.sizer.Add((1,5))

        # Add a checkbox for creating Quotes
        self.chkCreateQuotes = wx.CheckBox(self, -1, " " + _("Automatically Create Quotes"), style=wx.CHK_2STATE)
        # Check it by default
        self.chkCreateQuotes.SetValue(True)
        self.sizer.Add(self.chkCreateQuotes, 0, wx.ALL, 5)

        # Add a spacer
        self.sizer.Add((1,5))

        # Add the first Instruction Text
        instructions3 = _("Please select items for auto-coding.")
        txtInstructions3 = wx.StaticText(self, -1, instructions3)
        self.sizer.Add(txtInstructions3, 0, wx.ALL, 5)

        # Add a ListBox that allows selection of what items to auto-code, with multi-select enabled.
        self.autocode = wx.ListBox(self, -1, style = wx.LB_EXTENDED)
        self.sizer.Add(self.autocode, 1, wx.EXPAND | wx.ALL, 5)
        # Add a handler for auto-code item selection
        self.Bind(wx.EVT_LISTBOX, self.OnItemSelect)

    def IsComplete(self):
        """ Selections are NOT required on this page, so it's always complete. """
        return True

    def OnItemSelect(self, event):
        """ Handler for Auto-Code Item Selection """
        # Check Next Button Enabling on each change
        self.nextButton.Enable(self.IsComplete())


class SpreadsheetDataImport(wiz.Wizard):
    """ This displays the main Spreadsheet Data Import Wizard window. """
    def __init__(self, parent, treeCtrl):
        """ Initialize the Spreadsheet Data Import Wizard """
        # Remember the TreeCtrl
        self.treeCtrl = treeCtrl

        # Initialize data for the Wizard
        # This list holds the data imported from the file in a list of lists
        self.all_data = []
        # This list holds the Questions / Prompts from the data file
        self.questionsFromData = []
        # This dictionary holds Keyword Groups (keys) and Keywords (data) for Auto-Codes
        self.all_codes = {}

        # Create the Wizard
        wiz.Wizard.__init__(self, parent, -1, _('Import Spreadsheet Data'))
        self.SetPageSize(wx.Size(600, 450))

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Define the individual Wizard Pages
        self.FileNamePage = GetFileNamePage(self, _("Select a Spreadsheet Data File"))
        self.RowsOrColumnsPage = GetRowsOrColumnsPage(self, _("Identify Data Orientation"))
        self.ItemsToIncludePage = GetItemsToIncludePage(self, _("Select Data for Import"))
        self.AutoCodePage = GetAutoCodePage(self, _("Select Items for Auto-Coding."))

        # Define the page order / relationships
        self.FileNamePage.SetNext(self.RowsOrColumnsPage)
        self.RowsOrColumnsPage.SetPrev(self.FileNamePage)
        self.RowsOrColumnsPage.SetNext(self.ItemsToIncludePage)
        self.ItemsToIncludePage.SetPrev(self.RowsOrColumnsPage)
        self.ItemsToIncludePage.SetNext(self.AutoCodePage)
        self.AutoCodePage.SetPrev(self.ItemsToIncludePage)

        # Bind Wizard Events
        self.Bind(wiz.EVT_WIZARD_PAGE_CHANGING, self.OnPageChanging)
        self.Bind(wiz.EVT_WIZARD_PAGE_CHANGED, self.OnPageChanged)
        self.Bind(wiz.EVT_WIZARD_FINISHED, self.OnWizardFinished)

        # We need to add a HELP button to the Wizard Object.
        # First, let's create a Help Button
        self.helpButton = wx.Button(self, -1, _("Help"))
        self.helpButton.Bind(wx.EVT_BUTTON, self.OnHelp)

        # We need to figure out where to insert it into the existing Wizard Infrastructure.  It should go into the LAST
        # sizer we find in the Wizard object.  Let's initialize a variable so we can seek that sizer out.
        lastSizer = None
        # Sizers can be nested inside sizers.  So we iterate through the Wizard's Sizer's Children's sizer's Children to find it!
        for x in self.GetSizer().GetChildren()[0].Sizer.GetChildren():
            # If the current child is a sizer ...
            if x.IsSizer():
                # ... it's a candidate for the LAST sizer.  Remember it.
                lastSizer = x.Sizer
        # Add the Help button to the last sizer found!!
        if lastSizer != None:
            lastSizer.Add(self.helpButton, 0, wx.ALL, 5)

        # Run the Wixard
        self.FitToPage(self.FileNamePage)
        self.RunWizard(self.FileNamePage)

    def strip_quotes(self, text):
        """ Process text items from a Streadsheet Data File to strip leading and trailing quotes and remove double and triple quotes """
        # Remove leading and trailing whitespace
        text = text.strip()
        # Replace triple quotes with single quotes
        text = text.replace('"""', '"')
        # Replace double quotes with single quotes
        text = text.replace('""', '"')
        # If we have text that starts and ends with quotes ...
        if (len(text) > 1) and (text[0] == '"') and (text[-1] == '"'):
            # Strip a quote from the first and last positions in the string
            text = text[1:]
            text = text[:-1]
        # Return the processed string
        return text

    def fix_chars(self, text):
        """ Convert from HTML-ized special characters to regular special characters """
        # Clean up the text
        text = text.replace('&amp;', '&')
        text = text.replace('&gt;', '>')
        text = text.replace('&lt;', '<')
        return text

    def unfix_chars(self, text):
        """ Convert from regular special characters to HTML-ized special characters """
        # Clean up the text
        text = text.replace('&', '&amp;')
        text = text.replace('>', '&gt;')
        text = text.replace('<', '&lt;')
        return text

    def OnPageChanging(self, event):
        """ Process Wizard Page changes, with VETO option """
        if event.GetPage() == self.FileNamePage:
            prompt = unicode(_('File:  %s'), 'utf8')
            # ... set the File Name to the file selected on the File Name Page
            self.RowsOrColumnsPage.fileName.SetLabel(prompt % self.FileNamePage.txtSrcFileName.GetValue())
            # If we're moving FORWARD ...
            if event.GetDirection():
                # Initialize the File Data list
                self.all_data = []
                # Count number of row elements based on header
                rowElements = 0
                # Note the filename
                filename = self.FileNamePage.txtSrcFileName.GetValue()
                # Get the Character Encoding
                encoding = self.FileNamePage.txtSrcEncoding.GetStringSelection()
                # Open the file
                with open(filename, 'r') as f:
                    try:

                        ##########################################################################################
                        #                                                                                        #
                        #  NOTE:  The python csv reader has some problems.  For example, the following data:     #
                        #                                                                                        #
                        #         There is a "problem", at all times, handling a line with quotes and commas.    #
                        #                                                                                        #
                        #         gets parsed as 3 elements instead of one, and may get double-quotes at the     #
                        #         end of the word "problem"".  My code detects and adjusts for that.             #
                        #                                                                                        #
                        ##########################################################################################
                        
                        # Use the csv Sniffer to determine the dialect of the file
                        dialect = csv.Sniffer().sniff(f.read(1024))
                        # Reset the file to the beginning
                        f.seek(0)

                        # NOTE:  Tab delimited TXT files cause an IndexError here.  I thought the csvreader was supposed to
                        #        be able to handle tab delimited text.  I guess not.
                        
                        # use the csv Reader to read the data file
                        csvReader = csv.reader(f, dialect=dialect)
                        # For each row of data read ...
                        for row in csvReader:
                            # ... create a list for the Unicode encoded data
                            encRow = []
                            # Initialize variables for handling "problem" data.
                            # These variables allow us to combine data elements and know that we are doing it.
                            combine = False
                            combinedElement = ''
                            # For each element in the row  ...
                            for element in row:
                                # Process certain problem characters
                                element = self.unfix_chars(element)
                                # A weird unicode version of apostrophe that crept into one of my data files
                                element = element.replace(chr(146), chr(39))

                                # If we ARE handling weird data, or if we have an element with a Quotation Mark ...
                                if combine or (element.find('"') > -1):
                                    # If the LAST character is a quotation mark, that signals the end of the combining data.
                                    if element[-1] == '"':
                                        # If this is not a stand-alone (single) line, add the comma that was removed from the text
                                        if combinedElement != '':
                                            combinedElement += u", "
                                        # ... and add the combined elements as a single text element, dropping the end quotation mark
                                        encRow.append(combinedElement + element[:-1])
                                        # Reset the variables indicating that we are combining elements, as we're done.
                                        combinedElement = ''
                                        combine = False
                                    # If this is NOT a Final Character Quote situation ...
                                    else:
                                        # If this is not a stand-alone (single) line, add the comma that was removed from the text
                                        if combinedElement != '':
                                            combinedElement += u", "
                                        # Add the new element text to the combined element string
                                        combinedElement += element
                                        # Signal that we're in the middle of combining elements
                                        combine = True
                                # If we're NOT handling weird data and we don't have quotation marks in our text ...
                                else:
                                    # ... just add the element to the data structure
                                    encRow.append(element)

                            # ... add that encoded row to the data list
                            self.all_data.append(encRow)

                    except UnicodeDecodeError:
                        prompt = _("Unicode Error.  The encoding you selected does not match the data file you selected.")
                        tmpDlg = Dialogs.ErrorDialog(self, prompt)
                        tmpDlg.ShowModal()
                        tmpDlg.Destroy()
                        event.Veto()
                        return
                    
                    # If we get a CSV error, an IndexError, or a TypeError, let's try to do the import the hard way.
                    # (This is most often because we have a Tab Delimited Text file which CSV apparently can't handle!)
                    except (csv.Error, IndexError, TypeError):

                        if DEBUG:
                            print "SpreadsheetDataImport EXCEPTION:", sys.exc_info()[0], sys.exc_info()[1]

                        # Reset the data list.  We're starting over.                        
                        self.all_data = []
                        # Count number of row elements based on header
                        rowElements = 0
                        # Begin exception handling
                        try:
                            # Seek the start of the file
                            f.seek(0)
                            # Read the file's data
                            data = f.read()
                            # Decode the data based on the user-supplied file encoding
                            data = data.decode(encoding)
                            # If the last character of data is a newline character ...
                            if data[-1] == '\n':
                                # ... drop that newline character from data
                                data = data[:-1]
                            # Divide the data up into Rows.
                            rows = data.split('\n')
                            
                            # The problem with this method is that if a single answer contains a '\n' character within it,
                            # the whole data structue gets thrown off.  There are not enough tabs (columns) per row.
                            # We need to correct for this.

                            # Let's start by figuring out how many elements there should be, by looking at the first line.
                            # This could be problematic if data is organized in rows per participant and the first (0th)
                            # participant has this problem.
                            rowElements = rows[0].count('\t')
                            # Let's count the number of rows we're working with.
                            numRows = len(rows)
                            # We need to read the data structure from the bottom up because we are altering it as we go!
                            for x in range(numRows - 2, 0, -1):
                                # If our current row has more elements than the first row, we have a problem.
                                if rows[x].count('\t') > rowElements:

                                    # For now, just raise an exception.  Eventually, we may wish to add an error message or
                                    # even correct the problem!!
                                    raise(TransanaExceptions.NotImplementedError)

                                # If our current row has fewer elements than the first row ...
                                elif rows[x].count('\t') < rowElements:
                                    # ... add the current row to the previous row ...
                                    rows[x-1] += rows[x]
                                    # ... and remove the current row from the rows data structure
                                    rows = rows[:x] + rows[x+1:]
                            # The number of rows may have changed!
                            numRows = len(rows)
                            # for each row ...
                            for row in rows:
                                # ... initialize a list for row elements
                                row_data = []
                                # Break data row up at tabs
                                for element in row.split('\t'):
                                    # Process certain problem characters
                                    element = self.unfix_chars(element)
                                    # A weird unicode version of apostrophe that crept into one of my data files
                                    element = element.replace(u'\u2019', unichr(39))
                                    # ... append it to the row data list
                                    row_data.append(element)
                                # Now append the row data to the full data structure.
                                self.all_data.append(row_data)
                        except UnicodeDecodeError:
                            prompt = _("Unicode Error.  The encoding you selected does not match the data file you selected.")
                            tmpDlg = Dialogs.ErrorDialog(self, prompt)
                            tmpDlg.ShowModal()
                            tmpDlg.Destroy()
                            event.Veto()
                            return
                    except:

                        print sys.exc_info()[0]
                        print sys.exc_info()[1]

                        # Pass the exception on up the hierarchy
                        raise()

                # Place the items from the first row (that is,the first ROW of data) in the Row TextCtrl
                self.RowsOrColumnsPage.txt1.Clear()
                for col in self.all_data[0]:
                    self.RowsOrColumnsPage.txt1.AppendText(self.fix_chars(self.strip_quotes(col)) + u'\n')
                # Place the first item in each row (that is, the first COLUMN of data) in the Column TextCtrl
                self.RowsOrColumnsPage.txt2.Clear()
                for row in self.all_data:
                    self.RowsOrColumnsPage.txt2.AppendText(self.fix_chars(self.strip_quotes(row[0])) + u'\n')
                # Disable the Column and Row TextCtrls
                self.RowsOrColumnsPage.txt1.Enable(False)
                self.RowsOrColumnsPage.txt2.Enable(False)
        # If we move to the Organize and Include Items page ...
        elif event.GetPage() == self.RowsOrColumnsPage:
            # If we're moving FORWARD ...
            if event.GetDirection():
                # Initialize the Questions list
                self.questionsFromData = []
                # Determine the Questions if the user selected Columns ...
                if self.RowsOrColumnsPage.chkColumns.GetValue():
                    # ... by iterating through each row of data ...
                    for row in self.all_data:
                        # ... and selecting the row's first item
                        self.questionsFromData.append(self.fix_chars(self.strip_quotes(row[0])))
                # Determine the Questions if the user selected Rows ...
                else:
                    # ... by iterating through the first row of data
                    for col in self.all_data[0]:
                        # ... and selecting the column's header
                        self.questionsFromData.append(self.fix_chars(self.strip_quotes(col)))

                # Populate the combo of Questions / Prompts used to select the Unique Identifier after adding an automatic creation option
                self.ItemsToIncludePage.identifier.SetItems([_('Create one automatically')] + self.questionsFromData)
                self.ItemsToIncludePage.identifier.SetSelection(0)
                # Populate the list of Questions / Prompts
                self.ItemsToIncludePage.questions.SetItems(self.questionsFromData)
                # Select all items.  Select from the bottom up, including the first item.  (Stop at -1 to include item 0!)  This
                # means the list will be scrolled to the TOP.  
                for x in range(len(self.questionsFromData) - 1, -1, -1):
                    self.ItemsToIncludePage.questions.Select(x)
                # Scroll to the top
                # self.ItemsToIncludePage.questions.EnsureVisible(0)   # DOESN'T WORK
                # self.ItemsToIncludePage.questions.SetScrollPos(0, 0) # DOESN'T WORK

        # If we move to the Auto-Code Page ...
        elif event.GetPage() == self.ItemsToIncludePage:
            # If we're moving FORWARD ...
            if event.GetDirection():
                # ... set the Auto-Code options to match the Questions
                self.AutoCodePage.autocode.SetItems(self.questionsFromData)

    def OnPageChanged(self, event):
        """ Process Wizard Page changes with no veto option """
        # If we move BACKWARDS to the File Name Page ...
        if event.GetPage() == self.FileNamePage:
            # Reset the Questions and Codes
            self.questionsFromData = []
            self.all_codes = {}

        # Identify the Next button
        nextButton = self.FindWindowById(wx.ID_FORWARD)
        # Enable (or not) the Next Button depending the new page's "completeness"
        nextButton.Enable(event.GetPage().IsComplete())

    def PosAdjust(self, text):
        """ We need a way to adjust Quote Position information for the HTMLification of special characters. """
        # Initialize to no adjustment at all.
        adjustment = 0
        # Define character strings to look for and amount to adjust for each
        characters = (('&amp;', 4), ('&gt;', 3), ('&lt;', 3))

        # NOTE:  We can't just use string.replace() because we need to know how many adjustments we made.
        # For each character string / adjustment pair ...
        for (character, adjAmt) in characters:
            # Find the first instance of the character string
            startPos = text.find(character)
            # If we find (another) one ...
            while startPos > -1:
                # ... increase the adjustment value by the indicated amount ...
                adjustment += adjAmt
                # ... and search for the next one, starting after the one we just found
                startPos = text.find(character, startPos + adjAmt)
        # Return the adjustment value
        return adjustment

    def OnWizardFinished(self, event):
        """ Process the data when the Wizard is completed """
        # Determine if we're supposed to create Quotes or not
        createQuote = self.AutoCodePage.chkCreateQuotes.GetValue()
        # The Wizard is triggered on one of the Tree's Library nodes.  Get the Record Number for this node.
        libraryNumber = self.treeCtrl.GetPyData(self.treeCtrl.GetSelections()[0]).recNum
        # Get the full Library record
        library = Library.Library(libraryNumber)
        # If we're creating Quotes ...
        if createQuote:
            # ... start exception handling 
            try:
                # Try to open the Collection that shares its name with the Library.
                collection1 = Collection.Collection(library.id, 0)
            # If a RecordNotFoundError is raised, the Collection doesn't exist ...
            except TransanaExceptions.RecordNotFoundError:
                # ... so we need to create it.
                collection1 = Collection.Collection()
                collection1.id = library.id
                collection1.parent = 0
                collection1.comment = unicode(_('Created during Spreadsheet Data Import for file "%s."'), 'utf8') % self.FileNamePage.txtSrcFileName.GetValue()
                collection1.keyword_group = unicode(_('Auto-code'), 'utf8')
                collection1.db_save()
            # Signal that collections need to be added to the Database Tree
            collectionsDisplayed = False

        # Get the Unique Identifier selected by the user
        id_col = self.ItemsToIncludePage.identifier.GetSelection() - 1

        # Initialize the Participant Counter to 1 (the first participant) rather than 0.
        participantCount = 1

        # Get the Character Encoding
        encoding = self.FileNamePage.txtSrcEncoding.GetStringSelection()

        # We need to move through the data differently if the source file is organized by Columns or by Rows.

        # ... initialize the auto-codes found for THIS participant
        codes = {}
        # Initialize the dictionary for Collections based on Questions
        collections = {}

        # If data is in Columns ...
        if self.RowsOrColumnsPage.chkColumns.GetValue():
            # ... then the number of participants is the number of elements per row
            numParticipants = len(self.all_data[0])

            if DEBUG:
                print
                print "Prompts in Columns", ' Participants:', len(self.all_data[0]) - 1, '  Questions:', len(self.all_data) - 1

        # If data is in Rows ...
        elif self.RowsOrColumnsPage.chkRows.GetValue():
            # ... then the number of participants is the number of rows in the data
            numParticipants = len(self.all_data)

            if DEBUG:
                print "Prompts in Rows", ' Participants:', len(self.all_data), '  Questions:', len(self.all_data[0])

        # Note the number of questions in this spreadsheet file
        numQuestions = len(self.ItemsToIncludePage.questions.GetSelections())
        # Create a Progress Dialog to display progress through the questions for each participant
        progressDlg = DualProgressBar(self, _('Spreadsheet Data Import'), _('Participants'), _('Questions'), numParticipants, numQuestions)

        # For each ROW ...
        for x in range(1, numParticipants):
            # If a "Participant Identifier" column has been selected ...
            if id_col != -1:
                # If data is in Columns ...
                if self.RowsOrColumnsPage.chkColumns.GetValue():
                    # ... then note the selected participant identifier
                    partID = self.all_data[id_col][x]
                # If data is in Rows ...
                elif self.RowsOrColumnsPage.chkRows.GetValue():
                    # ... then note the selected participant identifier
                    partID = self.all_data[x][id_col]
            # If the user requested automatic unique Participant IDs ...
            if (id_col == -1) or (self.strip_quotes(partID) == ''):
                # ... create a unique Participant ID and increment the Participant Counter
                participantID = unicode(_('Participant %04d'), 'utf8') % participantCount
                participantCount += 1
            # Otherwise, use the data the user requested
            else:
                participantID = self.strip_quotes(partID)

            if DEBUG:
                print "Participant %d of %d" % (participantCount - 1, numParticipants - 1)

            # Update the first (participants) progress indicator
            progressDlg.Update(1, x)

            # Initialize Quote Position, Question Number, and list to hold new Quotes prior to saving
            quotePos = 0
            questionNum = 0
            quoteList = []

            # Create Document by participantID
            tmpDoc = Document.Document()
            # Populate essential Document Properties
            tmpDoc.id = self.fix_chars(participantID)
            tmpDoc.library_num = libraryNumber
            tmpDoc.imported_file = self.FileNamePage.txtSrcFileName.GetValue()
            tmpDoc.import_date = datetime.datetime.now().strftime('%Y-%m-%d')
            # Initialize Document Text with the initial XML for a Transana-XML document
            tmpDoc.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
<paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10">
"""
            tmpDoc.plaintext = ''

            # For each Question that should be included in the output ...
            for q in sorted(self.ItemsToIncludePage.questions.GetSelections()):
                # ... populate the Document Text and Plain Text with Question and Response
                # ... extract the question and answer ...
                question = self.strip_quotes(self.questionsFromData[q])
                # If data is in Columns ...
                if self.RowsOrColumnsPage.chkColumns.GetValue():
                    # ... then select the correct cell for the question's response
                    answer = self.strip_quotes(self.all_data[q][x])
                # If data is in Rows ...
                elif self.RowsOrColumnsPage.chkRows.GetValue():
                    # ... then select the correct cell for the question's response
                    answer = self.strip_quotes(self.all_data[x][q])
                # Increment Question Number (Numbering questions each time)
                questionNum += 1

                if DEBUG:
                    print '  ', participantID, 'of', numParticipants, " Q%05d" % questionNum, 'of', numQuestions
                    print '    ', type(question), question.encode('utf8')
                    print '    ', type(answer), answer.encode('utf8')

                # ... populate the Document Text and Plain Text with Question and Response
                # Start with a Paragraph declaration
                tmpDoc.text += """    <paragraph leftindent="0" rightindent="0" parspacingafter="51">
<text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">"""
                # Add the Question
                tmpDoc.text += self.unfix_chars(question.encode('utf8')) # + '\n'
                # Close the Paragraph declaration
                tmpDoc.text += """</text>
</paragraph>
"""
                # Add the Question ot the Plain Text
                tmpDoc.plaintext += self.fix_chars(question) + u'\n'

                # An answer *might* contain multiple paragraphs separated by newlines.  If so, each
                # one needs its own paragraph specification.
                answers = answer.split('\n')
                # For each paragraph in an individual answer ...
                for answerPart in answers:
                    # ... make sure we don't just have an empty paragraph in the middle of several
                    if (answerPart.strip() != '') or (len(answers) == 1):
                        # Start with a Paragraph declaration, left indent of 1"
                        tmpDoc.text += """    <paragraph leftindent="254" rightindent="254" parspacingafter="51">
<text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">"""
                        # Add the Answer
                        tmpDoc.text += answerPart.encode('utf8') # + '\n'
                        # Close the Paragraph declaration
                        tmpDoc.text += """</text>
</paragraph>
"""
                        # Add the answer paragraph to the Plain Text
                        tmpDoc.plaintext += self.fix_chars(answerPart) + u'\n'

                # If we are creating Quotes ...
                if createQuote:
                    # If we already have a Collection for this Question in our list of Collections ...
                    if questionNum in collections.keys():
                        # ... use the existing Collection
                        collection2 = collections[questionNum]
                    # If we don't already have the Collection in our list of Collections ...
                    else:
                        # ... create a Collection ID based on Question Number and some question text ...
                        collectionID = unicode("Q%04d - %s", 'utf8') % (questionNum, self.fix_chars(question.strip()))
                        # Start exception handling
                        try:
                            # ... try to open the Collection
                            collection2 = Collection.Collection(collectionID, collection1.number)
                        # If the Collection does not yet exist, Transana will raise a RecordNotFoundError 
                        except TransanaExceptions.RecordNotFoundError:
                            # ... indicating that we need to create the Collection.
                            collection2 = Collection.Collection()
                            collection2.id = collectionID[:50]
                            collection2.parent = collection1.number
                            collection2.comment = unicode(_('Created during Spreadsheet Data Import for file "%s."'), 'utf8') % self.FileNamePage.txtSrcFileName.GetValue()
                            collection2.keyword_group = unicode(_('Auto-code'), 'utf8')
                            collection2.db_save()
                        # Once the collection has been loaded or created ...
                        finally:
                            # ... add it to the List of Collections
                            collections[questionNum] = collection2

                    # Determine the correct value for the sort order for THIS collection
                    sortOrder = DBInterface.getMaxSortOrder(collection2.number) + 1

                    # Create a new Quote
                    quote1 = Quote.Quote()
                    # Create a Quote ID from the Participant ID and the Question Number
                    quote1.id = self.fix_chars(participantID) + unicode("  Q%04d" % questionNum, 'utf8')
                    # Populate the Quote's Collection Number and Collection ID
                    quote1.collection_num = collection2.number
                    quote1.collection_id = collection2.id
                    # Add the Sort Order
                    quote1.sort_order = sortOrder
                    # Add the Quote Starting Point, which we shoudl be tracking across *Document* creation.
                    quote1.start_char = quotePos
                    # Add the length of the Question and the answer and the two newline characters to the cumulative Quote Position ...
                    quotePos += len(question) + len(answer) + 2
                    # ... but then adjust these values for special characters such as &, <, and >.
                    quotePos -= self.PosAdjust(question)
                    quotePos -= self.PosAdjust(answer)

                    # We can use the Clip Transcripts as Quote Text too!
                    quote1.text = """<?xml version="1.0" encoding="UTF-8"?>
<richtext version="1.0.0.0" xmlns="http://www.wxwidgets.org">
<paragraphlayout textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New" alignment="1" leftindent="0" leftsubindent="0" rightindent="0" parspacingafter="10" parspacingbefore="0" linespacing="10">
"""
                    quote1.text += """    <paragraph leftindent="0" rightindent="0" parspacingafter="51">
  <text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">"""
                    # Add the Question
                    quote1.text += self.unfix_chars(question.encode('utf8')) # + '\n'
                    # Close the Paragraph declaration
                    quote1.text += """</text>
</paragraph>
"""
                    # Add the Question ot the Plain Text
                    quote1.plaintext = self.fix_chars(question) + u'\n'

                    # An answer *might* contain multiple paragraphs separated by newlines.  If so, each
                    # one needs its own paragraph specification.
                    for answerPart in answers:
                        # ... make sure we don't just have an empty paragraph in the middle of several
                        if (answerPart.strip() != '') or (len(answers) == 1):
                            # Start with a Paragraph declaration, left indent of 1"
                            quote1.text += """    <paragraph leftindent="254" rightindent="254" parspacingafter="51">
<text textcolor="#000000" bgcolor="#FFFFFF" fontpointsize="12" fontstyle="90" fontweight="90" fontunderlined="0" fontface="Courier New">"""
                            # Add the Answer
                            quote1.text += answerPart.encode('utf8') # + '\n'
                            # Close the Paragraph declaration
                            quote1.text += """</text>
</paragraph>
"""
                            # Add the answer paragraph to the Plain Text
                            quote1.plaintext += self.fix_chars(answerPart) + u'\n'

                    # Complete the Transana-XML documeht specification
                    quote1.text += """  </paragraphlayout>
</richtext>
"""
                    # We've adjusted the Quote Position variable.  We use that (-1) for the Quote End Position character
                    quote1.end_char = quotePos - 1
                    # Add a comment noting how this Quote was created.
                    quote1.comment = unicode(_('Created during Spreadsheet Data Import for file "%s."'), 'utf8') % self.FileNamePage.txtSrcFileName.GetValue()
                    # Add the quote to the Quote List.  Quotes will be completed and saved later.
                    quoteList.append(quote1)

            # Complete the Transana-XML document specification
            tmpDoc.text += """  </paragraphlayout>
</richtext>
"""

            # Remove trailing carriage return
            tmpDoc.plaintext = tmpDoc.plaintext[:-1]

            # For each selected Auto-Code category ...
            for c in self.AutoCodePage.autocode.GetSelections():
                # Define the Keyword Group
                kwg = unicode(_('Auto-code'), 'utf8')
                # Define the Keyword
                # If data is in Columns ...
                if self.RowsOrColumnsPage.chkColumns.GetValue():
                    # ... then the keyword is here
                    kwData = self.all_data[c][x]
                # If data is in Rows ...
                elif self.RowsOrColumnsPage.chkRows.GetValue():
                    # ... then the keyword is here
                    kwData = self.all_data[x][c]
                kw = unicode("%s : %s", 'utf8') % (self.strip_quotes(self.questionsFromData[c]), self.strip_quotes(kwData))
                # Replace HTML-ized characters, as well as replacing Parentheses (illegal in Keywords) with Brackets
                kw = self.fix_chars(kw.replace('(', '['))
                kw = kw.replace(')', ']')
                # Adjust for length
                kwg = kwg[:45]
                kw = kw[:80]

                if DEBUG:
                    print '      ', kwg, kw
                
                # If there was no missing data in the Keyword Definition ...
                if (self.strip_quotes(self.questionsFromData[c]) != '') and (self.strip_quotes(kwData) != ''):
                    # ... Add the Keyword to the Document
                    tmpDoc.add_keyword(kwg, kw)
                    # If we're creating Quotes ...
                    if createQuote:
                        # ... iterate through the quotes in the Quote List ...
                        for quote in quoteList:
                            # ... and add the keyword to the Quote
                            quote.add_keyword(kwg, kw)
                        
                    # If the Keyword Group had not been defined ...
                    if not kwg in self.all_codes.keys():
                        # ... define the Keyword Group using a list containing the Keyword
                        self.all_codes[kwg] = [kw]
                        # Try to load the keyword to see if it already exists.
                        try:
                            keyword = KeywordObject.Keyword(kwg, kw)
                        # If the Keyword doesn't exist yet ...
                        except TransanaExceptions.RecordNotFoundError:
                            # ... create the Keyword.
                            keyword = KeywordObject.Keyword()
                            keyword.keywordGroup = kwg
                            keyword.keyword = kw
                            keyword.definition = unicode(_('Created during Spreadsheet Data Import for file "%s."'), 'utf8') % self.FileNamePage.txtSrcFileName.GetValue()
                            # Try to save the keyword
                            keyword.db_save()
                            # Add the new Keyword to the database tree
                            self.treeCtrl.add_Node('KeywordNode', (_('Keywords'), keyword.keywordGroup, keyword.keyword), 0, keyword.keywordGroup)

                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (keyword.keywordGroup, keyword.keyword))
                    # If the Keyword Group HAS been defined ...
                    else:
                        # ... if the Keyword has NOT been defined ...
                        if not kw in self.all_codes[kwg]:
                            # ... add the Keyword to the Keyword Group's Keyword List
                            self.all_codes[kwg].append(kw)
                            # Try to load the keyword to see if it already exists.
                            try:
                                keyword = KeywordObject.Keyword(kwg, kw)
                            # If the Keyword doesn't exist yet ...
                            except TransanaExceptions.RecordNotFoundError:
                                # ... create the Keyword.
                                keyword = KeywordObject.Keyword()
                                keyword.keywordGroup = kwg
                                keyword.keyword = kw
                                keyword.definition = unicode(_('Created during Spreadsheet Data Import for file "%s."'), 'utf8') % self.FileNamePage.txtSrcFileName.GetValue()
                                try:
                                    # Try to save the Keyword
                                    keyword.db_save()
                                    # Add the new Keyword to the database tree
                                    self.treeCtrl.add_Node('KeywordNode', (_('Keywords'), keyword.keywordGroup, keyword.keyword), 0, keyword.keywordGroup)

                                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                                    if not TransanaConstants.singleUserVersion:
                                        if TransanaGlobal.chatWindow != None:
                                            TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (keyword.keywordGroup, keyword.keyword))
                                except:
                                    if TransanaConstants.demoVersion:
                                        prompt = unicode(_('Problem saving Keyword %s.  The Demo version limits the number of Keywords you may create.'), 'utf8')
                                        errorDlg = Dialogs.ErrorDialog(self, prompt % keyword.keyword)
                                        errorDlg.ShowModal()
                                        errorDlg.Destroy()

            # Prepare for a SaveError, just in case.  (The "finally" clause makes this necessary.)
            saveError = False
            # Try to save the new Document
            try:
                tmpDoc.db_save()
            # If there is a SaveError (duplicate name is most likely) ...
            except TransanaExceptions.SaveError:
                # Report the error to the user
                if TransanaConstants.demoVersion:
                    prompt = unicode(_('Problem saving Document %s.  The Demo version limits the number of Documents you may create.'), 'utf8')
                else:
                    prompt = unicode(_('Problem saving Document %s.  Try using an empty Collection.'), 'utf8')
                errorDlg = Dialogs.ErrorDialog(self, prompt % participantID)
                errorDlg.ShowModal()
                errorDlg.Destroy()
                # Note that we have a SaveError
                saveError = True
            
            finally:
                # If we had a SaveError ...
                if saveError:
                    # ... exit this module
                    return
                
                # Add the new Document to the Database Tree
                nodeData = (_('Libraries'), library.id, tmpDoc.id)
                self.treeCtrl.add_Node('DocumentNode', nodeData, tmpDoc.number, library.number)

                # Now let's communicate with other Transana instances if we're in Multi-user mode
                if not TransanaConstants.singleUserVersion:
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("AD %s >|< %s" % (nodeData[-2], nodeData[-1]))

                # If we're creating Quotes ...
                if createQuote:
                    # ... and the Collections have not already been displayed ...
                    if not collectionsDisplayed:
                        # Get all items for the Collection and add them to the tree using add_Node!
                        # Add the new Collection to the Database Tree
                        nodeData = (_('Collections'), collection1.id)
                        self.treeCtrl.add_Node('CollectionNode', nodeData, collection1.number, 0)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("AC %s" % collection1.id)

                        # For each Question Collection in the List of Collections ...
                        for collectionKey in collections.keys():
                            # ... get the Collection record ...
                            collection = collections[collectionKey]
                            # ... and add the new Collection to the Database Tree
                            nodeData = (_('Collections'), collection1.id, collection.id)
                            self.treeCtrl.add_Node('CollectionNode', nodeData, collection2.number, collection2.parent)

                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("AC %s >|< %s" % (nodeData[-2], nodeData[-1]))

                        # We have not added the collections.  We don't need to do that again.
                        collectionsDisplayed = True

                    # Initialize a Quote Counter (for Progress Notification only)
                    quoteNum = 0
                    # For each Quote in the Quote List ...
                    for quote in quoteList:
                        # ... increment the Quote Counter
                        quoteNum += 1

                        if DEBUG:
                            print "  Quote %d of %d" % (quoteNum, numQuestions)
                        # Update the second (questions) progress indicator
                        progressDlg.Update(2, quoteNum)

                        # Add the Source Document Number to the Quote, now that the Document Number is known
                        quote.source_document_num = tmpDoc.number
                        # Start exception handling
                        try:
                            # Save the Quote
                            quote.db_save()
                        except TransanaExceptions.SaveError:
                            if TransanaConstants.demoVersion:
                                prompt = unicode(_('Problem saving Quote %s for %s.  The Demo version limits the number of Quotes you may create.'), 'utf8')
                            else:
                                prompt = unicode(_('Problem saving Question %s for %s.'), 'utf8')
                            errorDlg = Dialogs.ErrorDialog(self, prompt % (quoteNum, participantID))
                            errorDlg.ShowModal()
                            errorDlg.Destroy()

                            return

                        # Add the new Quote to the Database Tree
                        nodeData = (_('Collections'),) + collection1.GetNodeData() + (quote.collection_id, quote.id, )
                        self.treeCtrl.add_Node('QuoteNode', nodeData, quote.number, quote.collection_num, sortOrder=quote.sort_order, avoidRecursiveYields=True)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("AQ %s >|< %s >|< %s" % (collection1.id, quote.collection_id, quote.id))

                if DEBUG:
                    print

        # If there are auto-codes ...
        if len(self.all_codes) > 0:
            # Now let's sort the Keywords Node
            keywordsNode = self.treeCtrl.select_Node((_('Keywords'),), 'KeywordRootNode')
            self.treeCtrl.SortChildren(keywordsNode)
            # Now let's sort the Keywords Auto-Code Node
            keywordsNode = self.treeCtrl.select_Node((_('Keywords'), _('Auto-code')), 'KeywordGroupNode')
            self.treeCtrl.SortChildren(keywordsNode)
        # Destroy the Progress Dialog box
        progressDlg.Destroy()

    def OnHelp(self, evt):
        """ Method to use when the Help Button is pressed """
        # If the Menu Window is defined ...
        if TransanaGlobal.menuWindow != None:
            # ... call it's Help() method for THIS control.
            TransanaGlobal.menuWindow.ControlObject.Help('Import Spreadsheet Data Files')

class DualProgressBar(wx.Dialog):
    """ A Dialog with dual Progress Bars """
    def __init__(self, parent, title, prompt1, prompt2, numItems1, numItems2):
        """ A Progress Dialog with two progress bars """

        # Create a Dialog Box
        wx.Dialog.__init__(self, parent, -1, title, size=(600, 150))
        # Add the Main Sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Add a gauge based on the number of records to be handled
        txtPrompt1 = wx.StaticText(self, -1, prompt1)
        mainSizer.Add(txtPrompt1, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        self.gauge1 = wx.Gauge(self, -1, numItems1)
        mainSizer.Add(self.gauge1, 0, wx.EXPAND | wx.ALL, 5)
        # Add a second gauge based on the second number of records to be handled
        txtPrompt2 = wx.StaticText(self, -1, prompt2)
        mainSizer.Add(txtPrompt2, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        self.gauge2 = wx.Gauge(self, -1, numItems2)
        mainSizer.Add(self.gauge2, 0, wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT, 5)

        # Finalize the Dialog layout
        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        # Center the Dialog on the Screen
        TransanaGlobal.CenterOnPrimary(self)
        # Initialize the gauges
        self.gauge1.SetValue(0)
        self.gauge2.SetValue(0)
        # Show the Dialog
        self.Show()

    def Update(self, gauge, value):
        """ Update gauge Number "gauge" to "value" """
        # If we are updating the first gauge ...
        if gauge == 1:
            # .. then update the first gauge
            self.gauge1.SetValue(value)
            # ... and automatically reset the second gauge
            self.gauge2.SetValue(0)
        # Otherwise ...
        else:
            # ... updage the second gauge
            self.gauge2.SetValue(value)

        
