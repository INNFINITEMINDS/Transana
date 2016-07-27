# Copyright (C) 2002-2016 Spurgeon Woods LLC
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

"""This module implements the DatabaseTreeTab class for the Data Display
Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case <nacase@wisc.edu>'

import DBInterface
import Dialogs
import TransanaGlobal

import wx
from Library import *

class LibraryPropertiesForm(Dialogs.GenForm):
    """Form containing Library fields."""

    def __init__(self, parent, id, title, library_object):
        self.width = 400
        self.height = 210
        # Make the Keyword Edit List resizable by passing wx.RESIZE_BORDER style
        Dialogs.GenForm.__init__(self, parent, id, title, size=(self.width, self.height), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
                                 useSizers = True, HelpContext='Library Properties')

        self.obj = library_object

        # Create the form's main VERTICAL sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a HORIZONTAL sizer for the first row
        r1Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v1 = wx.BoxSizer(wx.VERTICAL)
        # Library ID
        id_edit = self.new_edit_box(_("Library ID"), v1, self.obj.id, maxLen=100)
        # Add the element to the row sizer
        r1Sizer.Add(v1, 1, wx.EXPAND)

        # Add a horizontal spacer to the row sizer        
        r1Sizer.Add((10, 0))

        # Create a VERTICAL sizer for the next element
        v2 = wx.BoxSizer(wx.VERTICAL)
        # Owner
        owner_edit = self.new_edit_box(_("Owner"), v2, self.obj.owner, maxLen=100)
        # Add the element to the row sizer
        r1Sizer.Add(v2, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r1Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r2Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v3 = wx.BoxSizer(wx.VERTICAL)
        # Comment
        comment_edit = self.new_edit_box(_("Comment"), v3, self.obj.comment, maxLen=255)
        # Add the element to the row sizer
        r2Sizer.Add(v3, 1, wx.EXPAND)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r2Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a HORIZONTAL sizer for the next row
        r3Sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create a VERTICAL sizer for the next element
        v4 = wx.BoxSizer(wx.VERTICAL)
        self.kw_groups = DBInterface.list_of_keyword_groups()
        # Default KW Group
        self.kwg_choice = self.new_choice_box(_("Default Keyword Group"), v4, [""] + self.kw_groups, 0)
        # Add the element to the row sizer
        r3Sizer.Add(v4, 1, wx.EXPAND)
        if (self.obj.keyword_group) and (self.kwg_choice.FindString(self.obj.keyword_group) != wx.NOT_FOUND):
            self.kwg_choice.SetStringSelection(self.obj.keyword_group)
        else:
            self.kwg_choice.SetSelection(0)

        # Add the row sizer to the main vertical sizer
        mainSizer.Add(r3Sizer, 0, wx.EXPAND)

        # Add a vertical spacer to the main sizer        
        mainSizer.Add((0, 10))

        # Create a sizer for the buttons
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Add the buttons
        self.create_buttons(sizer=btnSizer)
        # Add the button sizer to the main sizer
        mainSizer.Add(btnSizer, 0, wx.EXPAND)
        # If Mac ...
        if 'wxMac' in wx.PlatformInfo:
            # ... add a spacer to avoid control clipping
            mainSizer.Add((0, 2))

        # Set the PANEL's main sizer
        self.panel.SetSizer(mainSizer)
        # Tell the PANEL to auto-layout
        self.panel.SetAutoLayout(True)
        # Lay out the Panel
        self.panel.Layout()
        # Lay out the panel on the form
        self.Layout()
        # Resize the form to fit the contents
        self.Fit()

        # Get the new size of the form
        (width, height) = self.GetSizeTuple()
        # Reset the form's size to be at least the specified minimum width
        self.SetSize(wx.Size(max(self.width, width), height))
        # Define the minimum size for this dialog as the current size, and define height as unchangeable
        self.SetSizeHints(max(self.width, width), height, -1, height)
        # Center the form on screen
        TransanaGlobal.CenterOnPrimary(self)

        # Set focus to Library ID
        id_edit.SetFocus()

    def get_input(self):
        """Show the dialog and return the modified Library Object.  Result
        is None if user pressed the Cancel button."""
        d = Dialogs.GenForm.get_input(self)     # inherit parent method
        if d:
            self.obj.id = d[_('Library ID')]
            self.obj.owner = d[_('Owner')]
            self.obj.comment = d[_('Comment')]
            self.obj.keyword_group = self.kwg_choice.GetStringSelection()
        else:
            self.obj = None

        return self.obj
        
class AddLibraryDialog(LibraryPropertiesForm):
    """Dialog used when adding a new Library."""

    def __init__(self, parent, id):
        obj = Library()
        obj.owner = DBInterface.get_username()
        LibraryPropertiesForm.__init__(self, parent, id, _("Add Library"), obj)


class EditLibraryDialog(LibraryPropertiesForm):
    """Dialog used when editing Library properties."""

    def __init__(self, parent, id, library_object):
        LibraryPropertiesForm.__init__(self, parent, id, _("Library Properties"), library_object)
