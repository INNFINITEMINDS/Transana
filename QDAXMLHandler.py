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

"""This module implements the Qualitative Data Exchange Format XML Import using Python's xml.sax."""

__author__ = 'David Woods <dwoods@transana.com>'

DEBUG = False
if DEBUG:
    print "QDAXMLHandler DEBUG is ON!!"

# import Python's xml.sax module
import xml.sax

# import Transana's Dialogs
import Dialogs
# import Transana's Keyword module
import KeywordObject as Keyword


class QDAXMLHandler(xml.sax.ContentHandler):
    """ Create an XML Handler following Python's xml.sax model """
    def __init__(self):
        """ Initialize the xml.sax parsing process """
        # Initialize the Object Type and Current Object to None
        self.objType = None
        self.currentObj = None
        # Create a dictionary to track the GUIDs for Codes.  This is useful because QDA-XML codes don't necessary follow Transana's format
        # and may require some manipulation before being put in the database
        self.keywordsDict = {}
        
    def startElement(self, name, attrs):
        """ Following Python's xml.sax model, read an XML Element Header and its attributes """
        if DEBUG:
            print 'startElement:'
            print '  ', name

        # Remember the element name.  It may be needed in characters()
        self.element = name

        # Determine the QDA-XML Element Type
        
        # If we have a CODE element ...
        if self.element.upper() == u'CODE':
            # ... our Object Type is Keyword
            self.objType = 'Keyword'
            # ... we'll use a dictionary to import Code data, as it will most likely not conform to our Keyword object (yet).
            self.currentObj = {}
            # Default Codes to Not Codable and a blank Definition
            self.currentObj['isCodable'] = False
            self.currentObj['definition'] = u''

        # Process QDA-XML Attributes

        # Iterate through (key, value) pairs of the Attributes
        for (key, value) in attrs.items():
            if DEBUG:
                print '    ', key, ':', value.encode('utf8')

            # If the Attribute Key is COLOR ...
            if key.upper() == u'COLOR':
                # If we have a Dictionary-style object ...
                if self.objType in ['Keyword']:
                    # ... add the Attribute Value to the dictionary's appropriate Key
                    self.currentObj['colorDef'] = value
            # If the Attribute Key is DESCRIPTION ...
            elif key.upper() == u'DESCRIPTION':
                # If we have a Dictionary-style object ...
                if self.objType in ['Keyword']:
                    # ... add the Attribute Value to the dictionary's appropriate Key
                    self.currentObj['definition'] = value
            # If the Attribute Key is GUID ...
            elif key.upper() == u'GUID':
                # If we have a Dictionary-style object ...
                if self.objType in ['Keyword']:
                    # ... add the Attribute Value to the dictionary's appropriate Key
                    self.currentObj['guid'] = value
            # If the Attribute Key is ISCODABLE ...
            elif key.upper() == u'ISCODABLE':
                # If we have a Dictionary-style object ...
                if self.objType in ['Keyword']:
                    # ... add the Attribute Value to the dictionary's appropriate Key
                    self.currentObj['isCodable'] = (value == u'true')
            # If the Attribute Key is NAME ...
            elif key.upper() == u'NAME':
                # If we have a Dictionary-style object ...
                if self.objType in ['Keyword']:
                    # ... add the Attribute Value to the dictionary's appropriate Key
                    self.currentObj['name'] = value
            
        if DEBUG:
            print

    def characters(self, data):
        """ Following Python's xml.sax model, read the characters between an XML Header tag and an XML Closing tag """
        if DEBUG:
            print "characters:"
            print '  ', data.encode('utf8'), type(data), self.element
            print

        # If there is data to be read (we don't have an empty line or just whitespace) ...
        if data.strip() != '':
            # ... if we have a QDA:ISCHILDOFCODE element ...
            if self.element.upper() == u'QDA:ISCHILDOFCODE':
                # If we have a Keyword object ...
                if self.objType == 'Keyword':
                    # ... the current Keyword's Parent has just been identified.  Transana needs the Keyword data for this GUID as the Keyword Group
                    #     for the Current Object!
                    self.currentObj['parent'] = data

    def endElement(self, name):
        """ Following Python's xml.sax model, process an XML Element Closing tag, which often means to process the now-complete imported object """
        if DEBUG:
            print "endElement:"
            print '  ', name, type(name), self.objType
            print

        # If we don't have an Object Type or a Current Object or if the Current object is an empy Dictionary ...
        if (self.objType == None) or (self.currentObj == None) or (self.currentObj == {}):
            # ... then we have nothing to do here.
            pass
        # If we have a Keyword Object ...
        elif self.objType == 'Keyword':
            # ... if we have a defined GUID ...
            if self.currentObj.has_key('guid'):
                # ... add the Current Object (a Dictionary containing Keyword data) to the Keyword Dictionary, using the GUID as the key
                self.keywordsDict[self.currentObj['guid']] = self.currentObj
                # If we have a DEFINITION ...
                if self.currentObj.has_key('definition'):
                    # ... strip of trailing white space.  Allow for leading white space??
                    self.currentObj['definition'] = self.currentObj['definition'].rstrip()

                if DEBUG:                
                    print 'GUID:       ', self.currentObj['guid']
                    if self.currentObj.has_key('parent'):
                        print 'Parent GUID:', self.currentObj['parent']
                        if self.keywordsDict.has_key('parent'):
                            print 'Parent:    ', self.keywordsDict['parent']['name'].encode('utf8')
                        else:
                            print 'Parent Unknown!'
                    print 'Name:       ', self.currentObj['name'].encode('utf8')
                    if self.currentObj.has_key('definition') and (self.currentObj['definition'] != ''):
                        print 'Definition:'
                        print self.currentObj['definition'].encode('utf8')
                    if self.currentObj.has_key('colorDef'):
                        print 'Color:      ', self.currentObj['colorDef']
                    print 'Codable:    ', self.currentObj['isCodable']
                    print
                    print

        # If we've reached the end of the file, we need to do final data processing.
        if (name.upper() in ['CODEBOOK', 'QDA-XML']):
            
            if DEBUG:
                print
                print
                print
                print
                print "Final Processing:"
                
                print "Keywords:"

            # Although it's unlikely, let's assume that keywords will fit Transana's requirements to start.
            # We need to format the codes very differently if in Transana format or not.
            keywordsInTransanaFormat = True

            # Sort out Keywords.  Are they in Transana Format (Keyword Groups in Keywords) or not?  We need to know.
            for key in self.keywordsDict.keys():

                if DEBUG:
                    print '  ', key, self.keywordsDict[key]

                # If we have not yet ruled Transana Format out ...
                if keywordsInTransanaFormat:

                    # Keyword Groups (first-level Codes) are not directly codable in Transana
                    if not(self.keywordsDict[key].has_key('parent')) and (self.keywordsDict[key]['isCodable'] == '1'):
                        keywordsInTransanaFormat = False

                        if DEBUG:
                            print ' -- A', self.keywordsDict[key]['name'], self.keywordsDict[key]['parent'], self.keywordsDict[key]['isCodable']

                    # Transana allows exactly two levels of keyword hierarchy.  Look for something with more than 2 levels.
                    if (self.keywordsDict[key].has_key('parent')) and (self.keywordsDict[self.keywordsDict[key]['parent']].has_key('parent')):
                        keywordsInTransanaFormat = False

                        if DEBUG:
                            print ' -- B', self.keywordsDict[key]['name'], self.keywordsDict[key]['parent'], self.keywordsDict[self.keywordsDict[key]['parent']].has_key('parent')

                    # Second Level Codes must be codable in Transana
                    if (self.keywordsDict[key].has_key('parent')) and (self.keywordsDict[key]['isCodable'] == '0'):
                        keywordsInTransanaFormat = False

                        if DEBUG:
                            print ' -- C', self.keywordsDict[key]['name'], self.keywordsDict[key]['parent'], self.keywordsDict[key]['isCodable']

                # If we have ruled Transana format out ...
                else:
                    # ... we're done here.
                    break

            if DEBUG:
                print
                print "Transana Format:", keywordsInTransanaFormat
                print

            # If we're working in Transana Keyword Format (Keywords in Keyword Groups) ...
            if keywordsInTransanaFormat:

                if DEBUG:
                    print "Save Keywords in Transana Format"

                # Iterate through the keywords
                for key in self.keywordsDict.keys():

                     if DEBUG:
                         print key, '-->', self.keywordsDict[key]

                     # If we DON'T have a Keyword Group and if our entry is codable (which it always should be in this format),
                     # then we have a full blown Keyword.  Otherwise, we can safely ignore the record!
                     if (self.keywordsDict[key].has_key('parent')) and (self.keywordsDict[key]['isCodable']):
                         # ... create a keyword object
                         keywordObj = Keyword.Keyword()
                         # ... the PARENT's Name in the keyword dictionary is the Keyword Group
                         keywordObj.keywordGroup = self.keywordsDict[self.keywordsDict[key]['parent']]['name']
                         # ... the Name is the Keyword
                         keywordObj.keyword = self.keywordsDict[key]['name']
                         # If there's a description ...
                         if self.keywordsDict[key].has_key('definition'):
                             # ... that's the definition
                             keywordObj.definition = self.keywordsDict[key]['definition']
                         # if there's a Color definition ...
                         if self.keywordsDict[key].has_key('colorDef'):
                             # ... then we have a color definition
                             keywordObj.lineColorDef = self.keywordsDict[key]['colorDef']

                         if DEBUG:
                             print keywordObj.keywordGroup.encode('utf8')
                             print keywordObj.keyword.encode('utf8')
                             print keywordObj.definition.encode('utf8')
                             print keywordObj.lineColorDef.encode('utf8')
                             print

                         # Save the Keyword
                         keywordObj.db_save()

            # If we have Codes in Non-Transana Format ...
            else:

                if DEBUG:
                    print "Save Keywords in a non-Transana Format"

                # Iterate through the Keyword Dictionary
                for key in self.keywordsDict.keys():

                    if DEBUG:
                        print key, '-->', self.keywordsDict[key]['name'].encode('utf8')

                    # If we have a codable entry ...
                    if (self.keywordsDict[key]['isCodable']):
                        # Create a Keyword Object
                        keywordObj = Keyword.Keyword()
                        # Declare "Imported Codes" as the Keyword Group
                        keywordObj.keywordGroup = _('Imported Codes')
                        # The Name is the Keyword, at least to start
                        keywordObj.keyword = self.keywordsDict[key]['name']
                        # If there's a description ...
                        if self.keywordsDict[key].has_key('definition'):
                            # ... that's the definition
                            keywordObj.definition = self.keywordsDict[key]['definition']
                        # if there's a Color definition ...
                        if self.keywordsDict[key].has_key('colorDef'):
                            # ... then we have a color definition
                            keywordObj.lineColorDef = self.keywordsDict[key]['colorDef']

                        # To handle code nesting , note the current Key value
                        tmpKey = key
                        # Climb the hierarchy until we get to parent == 0
                        while (self.keywordsDict[tmpKey].has_key('parent')):
                            # Remember the parent as the new key 
                            tmpKey = self.keywordsDict[tmpKey]['parent']
                            # Add the Name to the FRONT of the keyword value, colons separating levels
                            keywordObj.keyword = self.keywordsDict[tmpKey]['name'] + ' : ' + keywordObj.keyword

                        # if our Keyword exceeds Transana's max Keyword length ...
                        if len(keywordObj.keyword) > 85:
                            # ... create and display an error message for the user
                            msg = _('A Keyword is too long to be imported.  It will be shortened from:\n\n"%s"\n\nto:\n\n"%s"')
                            msgData = (keywordObj.keyword, u'... ' + keywordObj.keyword[-81:])
                            errDlg = Dialogs.ErrorDialog(None, msg % msgData, _("Import Error"))
                            errDlg.ShowModal()
                            errDlg.Destroy()

                            # Front-truncate the keyword to its maximum allowable length
                            keywordObj.keyword = u'... ' + keywordObj.keyword[-81:]

                        if DEBUG:
                             print keywordObj.keywordGroup.encode('utf8')
                             print keywordObj.keyword.encode('utf8')
                             print keywordObj.definition.encode('utf8')
                             print keywordObj.lineColorDef.encode('utf8')
                             print

                        # Save the Keyword
                        keywordObj.db_save()

                if DEBUG:
                    print

##        # Since we're done with the current 
##        self.currentObj = None
