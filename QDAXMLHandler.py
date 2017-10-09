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
        # Initialize the Object Type to None
        self.objType = None
        # Current Object is a list to handle nested objects
        self.currentObj = []
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
            # Note that we append the new dictionary to the current object list to handle nesting of codes.
            self.currentObj.append({})
            # Default Codes to Not Codable and a blank Definition
            self.currentObj[-1]['isCodable'] = False
            self.currentObj[-1]['definition'] = u''
        elif self.element.upper() in ['SET', 'SETS']:
            self.objType = None

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
                    self.currentObj[-1]['colorDef'] = value
##            # If the Attribute Key is DESCRIPTION ...
##            elif key.upper() == u'DESCRIPTION':
##                # If we have a Dictionary-style object ...
##                if self.objType in ['Keyword']:
##                    # ... add the Attribute Value to the dictionary's appropriate Key
##                    self.currentObj[-1]['definition'] = value
            # If the Attribute Key is GUID ...
            elif key.upper() == u'GUID':
                # If we have a Dictionary-style object ...
                if self.objType in ['Keyword']:
                    # ... add the Attribute Value to the dictionary's appropriate Key
                    self.currentObj[-1]['guid'] = value
            # If the Attribute Key is ISCODABLE ...
            elif key.upper() == u'ISCODABLE':
                # If we have a Dictionary-style object ...
                if self.objType in ['Keyword']:
                    # ... add the Attribute Value to the dictionary's appropriate Key
                    self.currentObj[-1]['isCodable'] = (value == u'true')
            # If the Attribute Key is NAME ...
            elif key.upper() == u'NAME':
                # If we have a Dictionary-style object ...
                if self.objType in ['Keyword']:
                    # ... add the Attribute Value to the dictionary's appropriate Key
                    self.currentObj[-1]['name'] = value
            
        if DEBUG:
            print

    def characters(self, data):
        """ Following Python's xml.sax model, read the characters between an XML Header tag and an XML Closing tag """
        if DEBUG:
            print "characters:  (", self.objType, self.element, ")"
            print '  "%s"' % data.encode('utf8'), type(data), len(data),
            if len(data) == 1:
                print '-->', ord(data)
            else:
                print
            print

        # If there is data to be read (we don't have an empty line or just whitespace) ...
        if (len(self.currentObj) > 0) and (self.element.upper() == 'DESCRIPTION'):
            # Handle NewLine Characters
            if (len(data) == 1) and (ord(data) == 10):
                self.currentObj[-1]['definition'] += '\n'
            else:
                self.currentObj[-1]['definition'] += data

    def endElement(self, name):
        """ Following Python's xml.sax model, process an XML Element Closing tag, which often means to process the now-complete imported object """
        if DEBUG:
            print "endElement:  (", self.objType, self.element, ")"
            print '  ', name, type(name), len(self.currentObj), self.currentObj
            print

        # If we don't have an Object Type or a Current Object or if the Current object is an empy Dictionary ...
        if (self.objType == None) or (self.currentObj == []) or (self.currentObj == [{}]) or \
           (name.upper() in ['CODEBOOK', 'CODES', 'DESCRIPTION', 'MEMBERCODE', 'SET', 'SETS']):
            # ... then we remove the element type.
            self.element = ''
        # If we have a Keyword Object ...
        elif self.objType == 'Keyword':
            # ... and the closing tag of a Code element ...
            if name.upper() == 'CODE' and (len(self.currentObj) > 0):
                # ... get the last element from our Current Object (pop removes it from the list and returns it)
                tmpObj = self.currentObj.pop()
                # If there are still elements in the Current Object, the last element is this code's Parent!
                if len(self.currentObj) > 0:
                    tmpObj['parent'] = self.currentObj[-1]['guid']
                # Add this new Keyword dictionary object to our Keyword Dictionary
                self.keywordsDict[tmpObj['guid']] = tmpObj
                
                if DEBUG:                
                    print 'GUID:       ', tmpObj['guid']
                    if tmpObj.has_key('parent'):
                        print 'Parent GUID:', tmpObj['parent']
                        if self.keywordsDict.has_key('parent'):
                            print 'Parent:    ', self.keywordsDict['parent']['name'].encode('utf8')
                        else:
                            print 'Parent Unknown!'
                    print 'Name:       ', tmpObj['name'].encode('utf8')
                    if tmpObj.has_key('definition') and (tmpObj['definition'] != ''):
                        print 'Definition:'
                        print tmpObj['definition'].encode('utf8')
                    if tmpObj.has_key('colorDef'):
                        print 'Color:      ', tmpObj['colorDef']
                    print 'Codable:    ', tmpObj['isCodable']
                    print
                    print
            if name.upper() in ['CODE']:
                self.element = ''



            

        # If we've reached the end of the file, we need to do final data processing.
        if (name.upper() in ['CODEBOOK', 'QDA-XML']):
            
            if DEBUG:
                print
                print
                print
                print
                print "Final Processing:"
                
                print "Keywords:"

                for x in self.keywordsDict.keys():
                    print x, self.keywordsDict[x]
                    print
                    
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
                    if not(self.keywordsDict[key].has_key('parent')) and (self.keywordsDict[key]['isCodable']):
                        keywordsInTransanaFormat = False

                        if DEBUG:
                            print ' -- A', self.keywordsDict[key]['name'], self.keywordsDict[key]['isCodable']

                    # Transana allows exactly two levels of keyword hierarchy.  Look for something with more than 2 levels.
                    if (self.keywordsDict[key].has_key('parent')) and (self.keywordsDict[self.keywordsDict[key]['parent']].has_key('parent')):
                        keywordsInTransanaFormat = False

                        if DEBUG:
                            print ' -- B', self.keywordsDict[key]['name'], self.keywordsDict[key]['parent'], self.keywordsDict[self.keywordsDict[key]['parent']].has_key('parent')

                    # Second Level Codes must be codable in Transana
                    if (self.keywordsDict[key].has_key('parent')) and (not self.keywordsDict[key]['isCodable']):
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
