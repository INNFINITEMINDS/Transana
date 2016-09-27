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

""" This module handles Search Requests and all related processing. """

__author__ = 'David Woods <dwoods@transana.com>'

DEBUG = False
if DEBUG:
    print "ProcessSearch DEBUG is ON!!"
    import datetime

# Import wxPython
import wx

# Import the Transana Collection Object
import Collection
# Import the Transana Database Interface
import DBInterface
# import Transana's Dialog
import Dialogs
# Import the Transana Document Object
import Document
# Import the Transana Episode Object
import Episode
# Import the Transana Library Object
import Library
# Import the Transana Quote Object
import Quote
# Import the Transana Search Dialog Box
import SearchDialog
# import Transana's Constants
import TransanaConstants
# Import Transana's Globals
import TransanaGlobal
# Import Transana's Transcript object
import Transcript

# Import the Python String module
import string

class ProcessSearch(object):
    """ This class handles all processing related to Searching. """
    # searchName and searchTerms are used by unit_test_search
    def __init__(self, dbTree, searchCount, kwg=None, kw=None, searchName=None, searchTerms=None, searchScope=None):
        """ Initialize the ProcessSearch class.  The dbTree parameter accepts a wxTreeCtrl as the Database Tree where
            Search Results should be displayed.  The searchCount parameter accepts the number that should be included
            in the Default Search Title. Optional kwg (Keyword Group) and kw (Keyword) parameters implement Quick Search
            for the keyword specified. searchName, searchTerms, and searchScope are used for searches triggereed by
            the Word Frequency Report (and by unit_test_search) """

        # See if there are any records that need Plain Text extraction
        plainTextCount = DBInterface.CountItemsWithoutPlainText()
        # If there are ...
        if plainTextCount > 0:
            # ... import the Plain Text extractor (which cannot be imported above, at least not in it's alphabetic position)
            import PlainTextUpdate
            # Create the Plain Text extractor Dialog
            tmpDlg = PlainTextUpdate.PlainTextUpdate(None, plainTextCount)
            # Show the Dialog
            tmpDlg.Show()
            # Begin the conversion / extraction
            tmpDlg.OnConvert()
            # Clean up when done.
            tmpDlg.Close()
            tmpDlg.Destroy()

        # Note the Database Tree that accepts Search Results
        self.dbTree = dbTree
        # Set up empty lists for different object types, including everything unless otherwise required
        self.documentList = []
        self.transcriptList = []
        self.collectionList = []
        # If kwg and kw are None, we are doing a regular (full) search.
        if ((kwg == None) or (kw == None)) and (searchTerms == None):
            # Create the Search Dialog Box
            dlg = SearchDialog.SearchDialog(_("Search") + " %s" % searchCount)
            # Display the Search Dialog Box and record the Result
            result = dlg.ShowModal()
            # If the user selects OK ...
            if result == wx.ID_OK:
                # ... get the search name from the dialog
                searchName = dlg.searchName.GetValue().strip()
                # Search Name is required.  If it was eliminated, put it back!
                if searchName == '':
                    searchName = _("Search") + " %s" % searchCount

                # Get the Collections Tree from the Search Form
                collTree = dlg.ctcCollections
                # Get the Collections Tree's Root Node
                collNode = collTree.GetRootItem()
                # Get a list of all the Checked Collections in the Collections Tree
                self.collectionList = dlg.GetCollectionList(collTree, collNode, True)
                # We need to check to see if there are ANY collections that are NOT checked.
                # If there are NOT any un-checked Collections, we can completely ignore Collection Specification,
                # making the search simpler and presumably faster!!
                uncheckedCollectionList = dlg.GetCollectionList(collTree, collNode, False)
                # If there are NO unchecked Collections ...
                if len(uncheckedCollectionList) == 0:
                    # ... we can ignore the CollectionList Completely!!
                    self.collectionList = []
                # ... and get the search terms from the dialog
                searchTerms = dlg.searchQuery.GetValue().split('\n')
                # Get the includeDocuments info
                includeDocuments = dlg.includeDocuments.IsChecked()
                # Get the includeEpisodes info
                includeEpisodes = dlg.includeEpisodes.IsChecked()
                # Get the includeQuotes info
                includeQuotes = dlg.includeQuotes.IsChecked()
                # Get the includeClips info
                includeClips = dlg.includeClips.IsChecked()
                # Get the includeSnapshots info
                includeSnapshots = dlg.includeSnapshots.IsChecked()
            # Destroy the Search Dialog Box
            dlg.Destroy()
        # SearchTerms are passed in by the Word Frequency Report (with searchScope) and
        # during Unit_Test_Search (without searchScope)
        elif (searchTerms != None):
            # There's no dialog.  Just say the user said OK.
            result = wx.ID_OK
            # Call from unit_test_search, so we can hard-code these parameters
            if searchScope == None:
                # Include Episodes and Clips.
                includeEpisodes = True
                includeClips = True
                # If Pro, Lab, or MU, include Documents, Quotes, and Snapshots.  
                if TransanaConstants.proVersion:
                    includeDocuments = True
                    includeQuotes = True
                    includeSnapshots = True
                    for term in searchTerms:
                        if u'Item Text contains' in term:
                            includeSnapshots = False
                            break
                else:
                    includeDocuments = False
                    includeQuotes = False
                    includeSnapshots = False
                    
            # We need to figure out the scope of the Word Frequency Report that triggered this search
            else:
                # Get the Node Data for the triggering tree node
                itemData = self.dbTree.GetPyData(searchScope)
                # Library Root, Library, and Document nodes need to show Documents.
                if itemData.nodetype in ['LibraryRootNode', 'LibraryNode', 'DocumentNode']:
                    includeDocuments = True
                else:
                    includeDocuments = False
                # Library Root, Library, and Episode nodes need to show Episodes.
                if itemData.nodetype in ['LibraryRootNode', 'LibraryNode', 'EpisodeNode']:
                    includeEpisodes = True
                else:
                    includeEpisodes = False
                # Library Root, Library, Episode, and Transcript nodes need to show Documents.
                if itemData.nodetype in ['LibraryRootNode', 'LibraryNode', 'EpisodeNode', 'TranscriptNode']:
                    includeTranscripts = True
                else:
                    includeTranscripts = False
                # Collection Root and Collection nodes need to show Quotes
                if itemData.nodetype in ['CollectionsRootNode', 'CollectionNode']:
                    includeQuotes = True
                else:
                    includeQuotes = False
                # We never include snapshots with text search!
                includeSnapshots = False
                # Collection Root and Collection nodes need to show Clips
                if itemData.nodetype in ['CollectionsRootNode', 'CollectionNode']:
                    includeClips = True
                else:
                    includeClips = False

                # Determine what Libraries, Documents, Episodes, Transcripts, and Collections to include or exclude.
                # Leave the lists empty if not applicable to simplify the SQL.

                # If we have a LibraryRoot Node, do NOTHING because we want ALL Documents and Episode Transcripts
                # If we have a Library Node ...
                if itemData.nodetype in ['LibraryNode']:
                    # ... we need only the Documents and Episode Transcripts within that Library.
                    self.documentList = self.GetNodeList(self.dbTree, searchScope, 'DocumentNode')
                    self.transcriptList = self.GetNodeList(self.dbTree, searchScope, 'TranscriptNode')
                # If we have a Document Node ...
                elif itemData.nodetype in ['DocumentNode']:
                    # ... we need only the selected Document.
                    self.documentList = [(itemData.recNum, self.dbTree.GetItemText(searchScope))]
                # If we have an Episode Node ...
                elif itemData.nodetype in ['EpisodeNode']:
                    # ... we need all Transcripts for that Episode.
                    self.transcriptList = self.GetNodeList(self.dbTree, searchScope, 'TranscriptNode')
                # If we have a Transcript Node ...
                elif itemData.nodetype in ['TranscriptNode']:
                    # ... we need only the selected Transcript.
                    self.transcriptList = [(itemData.recNum, self.dbTree.GetItemText(searchScope))]
                # If we have a Collection Node ...
                elif itemData.nodetype in ['CollectionNode']:
                    # ... we need the selected Collection plus all nested collections.
                    #     (The selected collection doesn't get included by the recursive call!)
                    self.collectionList = [(itemData.recNum, self.dbTree.GetItemText(searchScope))] + \
                                          self.GetNodeList(self.dbTree, searchScope, 'CollectionNode')

        # if kwg and kw are passed in, we're doing a Quick Search
        else:
            # There's no dialog.  Just say the user said OK.
            result = wx.ID_OK
            # The Search Name is built from the kwg : kw combination
            searchName = "%s : %s" % (kwg, kw)
            # The Search Terms are just the keyword group and keyword passed in
            searchTerms = ["%s:%s" % (kwg, kw)]
            # Include Clips.  Do not include Documents or Episodes
            includeDocuments = False
            includeEpisodes = False
            includeClips = True
            # If Pro, Lab, or MU, include Quotes and Snapshots.  
            if TransanaConstants.proVersion:
                includeQuotes = True
                includeSnapshots = True
            else:
                includeQuotes = False
                includeSnapshots = False

        # If OK is pressed (or Quick Search), process the requested Search
        if result == wx.ID_OK:
            # Increment the Search Counter
            self.searchCount = searchCount + 1
            # The "Search" node itself is always item 0 in the node list
            searchNode = self.dbTree.select_Node((_("Search"),), 'SearchRootNode')
            # We need to collect a list of the named searches already done.
            namedSearches = []
            # Get the first child node from the Search root node
            (childNode, cookieVal) = self.dbTree.GetFirstChild(searchNode)
            # As long as there are child nodes ...
            while childNode.IsOk():
                # Add the node name to the named searches list ...
                namedSearches.append(self.dbTree.GetItemText(childNode))
                # ... and get the next child node
                (childNode, cookieVal) = self.dbTree.GetNextChild(childNode, cookieVal)
            # We need to give each search result a unique name.  So note the search count number
            nameIncrementValue = searchCount
            # As long as there's already a named search with the name we want to use ...
            while (searchName in namedSearches):
                # ... if this is our FIRST attempt ...
                if nameIncrementValue == searchCount:
                    # ... append the appropriate number on the end of the search name
                    searchName += unicode(_(' - Search %d'), 'utf8') % nameIncrementValue
                # ... if this is NOT our first attempt ...
                else:
                    # ... remove the previous number and add the appropriate next number to try
                    searchName = searchName[:searchName.rfind(' ')] + ' %d' % nameIncrementValue
                # Increment our counter by one.  We'll keep trying new numbers until we find one that works.
                nameIncrementValue += 1
            # As long as there's a search name (and there's no longer a way to eliminate it!
            if searchName != '':

                # Build the appropriate Queries based on the Search Query specified in the Search Dialog.
                # (This method parses the Natural Language Search Terms into queries for Episode Search
                #  Terms, for Clip Search Terms, and for Snapshot Search Terms, and includes the appropriate 
                #  Parameters to be used with the queries.  Parameters are not integrated into the queries 
                #  in order to allow for automatic processing of apostrophes and other text that could 
                #  otherwise interfere with the SQL execution.)
                (documentQuery, episodeQuery, quoteQuery, clipQuery, wholeSnapshotQuery, snapshotCodingQuery, params, textSearchItems) = \
                    self.BuildQueries(searchTerms)

                # Clip Searches with Text seem to take a long time.  Let's display a Popup if there's Text.
                if len(textSearchItems) > 0:
                    progressDialog = Dialogs.PopupDialog(None, _('Search'), _('Search in progress.  Please wait.'))

                # Add a Search Results Node to the Database Tree
                nodeListBase = [_("Search"), searchName]
                self.dbTree.add_Node('SearchResultsNode', nodeListBase, 0, 0, expandNode=True, textSearchItems = textSearchItems)

                # Get a Database Cursor
                dbCursor = DBInterface.get_db().cursor()

                if DEBUG:
                    t1 = datetime.datetime.now()

                # Episodes
                if includeEpisodes:
                    # Adjust query for sqlite, if needed
                    episodeQuery = DBInterface.FixQuery(episodeQuery)
                    # Execute the Library/Episode query
                    dbCursor.execute(episodeQuery, tuple(params))
                    # Process the results of the Library/Episode query
                    for line in DBInterface.fetchall_named(dbCursor):
                        # Add the new Transcript(s) to the Database Tree Tab.
                        # To add a Transcript, we need to build the node list for the tree's add_Node method to climb.
                        # We need to add the Library, Episode, and Transcripts to our Node List, so we'll start by loading
                        # the current Library and Episode
                        tempLibrary = Library.Library(line['SeriesNum'])
                        tempEpisode = Episode.Episode(line['EpisodeNum'])
                        # Add the Search Root Node, the Search Name, and the current Library and Episode Names.
                        nodeList = (_('Search'), searchName, tempLibrary.id, tempEpisode.id)
                        # In order to include Clips without Transcripts (when there is no Text Search element),
                        # line may or may not include a TranscriptNum dictionary element.  If it does, only include specified
                        # Transcripts in the search results.
                        if line.has_key('TranscriptNum'):
                            tempTranscript = Transcript.Transcript(line['TranscriptNum'])
                            nodeList += (tempTranscript.id,)
                            # Add the Transcript Node to the Tree.  
                            self.dbTree.add_Node('SearchTranscriptNode', nodeList, tempTranscript.number, tempTranscript.episode_num, textSearchItems = textSearchItems)
                        # If line does NOT include a TranscriptNum, load all available transcripts for the search results.
                        else:
                            # Find out what Transcripts exist for each Episode
                            transcriptList = DBInterface.list_transcripts(tempLibrary.id, tempEpisode.id)
                            # If the Episode HAS defined transcripts ...
                            if len(transcriptList) > 0:
                                # Add each Transcript to the Database Tree
                                for (transcriptNum, transcriptID, episodeNum) in transcriptList:
                                    # Add the Transcript Node to the Tree.  
                                    self.dbTree.add_Node('SearchTranscriptNode', nodeList + (transcriptID,), transcriptNum, episodeNum, textSearchItems = textSearchItems)
                            # If the Episode has no transcripts, it still has the keywords and SHOULD be displayed!
                            else:
                                # Add the Transcript-less Episode Node to the Tree.  
                                self.dbTree.add_Node('SearchEpisodeNode', nodeList, tempEpisode.number, tempLibrary.number, textSearchItems = textSearchItems)

                if DEBUG:
                    t2 = datetime.datetime.now()
                    
                # Documents
                if includeDocuments:
                    # Adjust query for sqlite, if needed
                    documentQuery = DBInterface.FixQuery(documentQuery)
                    # Execute the Library/Document query
                    dbCursor.execute(documentQuery, tuple(params))

                    # Process the results of the Library/Document query
                    for line in DBInterface.fetchall_named(dbCursor):
                        # Add the new Document(s) to the Database Tree Tab.
                        # To add a Document, we need to build the node list for the tree's add_Node method to climb.
                        # We need to add the Library and Documents to our Node List, so we'll start by loading
                        # the current Library
                        tempLibraryName = DBInterface.ProcessDBDataForUTF8Encoding(line['SeriesID'])
                        tempDocument = Document.Document(line['DocumentNum'])
                        # Add the Search Root Node, the Search Name, and the current Library Name.
                        nodeList = (_('Search'), searchName, tempLibraryName)
                        # Add the Document Node to the Tree.
                        self.dbTree.add_Node('SearchDocumentNode', nodeList + (tempDocument.id,), tempDocument.number, tempDocument.library_num, textSearchItems = textSearchItems)

                if DEBUG:
                    t3 = datetime.datetime.now()
                    
                # Quotes
                if includeQuotes:
                    # Adjust query for sqlite, if needed
                    quoteQuery = DBInterface.FixQuery(quoteQuery)
                    # Execute the Collection/Quote query
                    dbCursor.execute(quoteQuery, params)

                    # Process all results of the Collection/Quote query 
                    for line in DBInterface.fetchall_named(dbCursor):
                        # Add the new Quote to the Database Tree Tab.
                        # To add a Quote, we need to build the node list for the tree's add_Node method to climb.
                        # We need to add all of the Collection Parents to our Node List, so we'll start by loading
                        # the current Collection
                        tempCollection = Collection.Collection(line['CollectNum'])

                        # Add the current Collection Node Data
                        nodeList = tempCollection.GetNodeData()                        
                        # Get the DB Values
                        tempID = line['QuoteID']
                        # If we're in Unicode mode, format the strings appropriately
                        if 'unicode' in wx.PlatformInfo:
                            tempID = DBInterface.ProcessDBDataForUTF8Encoding(tempID)
                        # Now add the Search Root Node and the Search Name to the front of the Node List and the
                        # Quote Name to the back of the Node List
                        nodeList = (_('Search'), searchName) + nodeList + (tempID, )

                        # Add the Node to the Tree
                        self.dbTree.add_Node('SearchQuoteNode', nodeList, line['QuoteNum'], line['CollectNum'], sortOrder=line['SortOrder'], textSearchItems = textSearchItems)

                if DEBUG:
                    t4 = datetime.datetime.now()
                    
                # Clips
                if includeClips:
                    # Adjust query for sqlite, if needed
                    clipQuery = DBInterface.FixQuery(clipQuery)
                    # Execute the Collection/Clip query
                    dbCursor.execute(clipQuery, params)

                    # Process all results of the Collection/Clip query 
                    for line in DBInterface.fetchall_named(dbCursor):
                        # Add the new Clip to the Database Tree Tab.
                        # To add a Clip, we need to build the node list for the tree's add_Node method to climb.
                        # We need to add all of the Collection Parents to our Node List, so we'll start by loading
                        # the current Collection
                        tempCollection = Collection.Collection(line['CollectNum'])

                        # Add the current Collection Node Data
                        nodeList = tempCollection.GetNodeData()                        
                        # Get the DB Values
                        tempID = line['ClipID']
                        # If we're in Unicode mode, format the strings appropriately
                        if 'unicode' in wx.PlatformInfo:
                            tempID = DBInterface.ProcessDBDataForUTF8Encoding(tempID)
                        # Now add the Search Root Node and the Search Name to the front of the Node List and the
                        # Clip Name to the back of the Node List
                        nodeList = (_('Search'), searchName) + nodeList + (tempID, )

                        # Add the Node to the Tree
                        self.dbTree.add_Node('SearchClipNode', nodeList, line['ClipNum'], line['CollectNum'], sortOrder=line['SortOrder'], textSearchItems = textSearchItems)

                if DEBUG:
                    t5 = datetime.datetime.now()

                    print "Episodes: ", t2 - t1
                    print "Documents: ", t3 - t2
                    print "Quotes: ", t4 - t3
                    print "Clips: ", t5 - t4
                    print "Total: ", t5 - t1
                    
                # If Snapshots are check AND there is no Text Search Component ...
                # (If there is a Text Search component to the search, the wholeSnapshotQuery is blank!!)
                if includeSnapshots and wholeSnapshotQuery != '':
                    # Adjust query for sqlite, if needed
                    wholeSnapshotQuery = DBInterface.FixQuery(wholeSnapshotQuery)
                    # Execute the Whole Snapshot query
                    dbCursor.execute(wholeSnapshotQuery, params)

                    # Since we have two sources of Snapshots that get included, we need to track what we've already
                    # added so we don't add the same Snapshot twice
                    addedSnapshots = []

                    # Process all results of the Whole Snapshot query 
                    for line in DBInterface.fetchall_named(dbCursor):
                        # Add the new Snapshot to the Database Tree Tab.
                        # To add a Snapshot, we need to build the node list for the tree's add_Node method to climb.
                        # We need to add all of the Collection Parents to our Node List, so we'll start by loading
                        # the current Collection
                        tempCollection = Collection.Collection(line['CollectNum'])

                        # Add the current Collection Node Data
                        nodeList = tempCollection.GetNodeData()                        
                        # Get the DB Values
                        tempID = line['SnapshotID']
                        # If we're in Unicode mode, format the strings appropriately
                        if 'unicode' in wx.PlatformInfo:
                            tempID = DBInterface.ProcessDBDataForUTF8Encoding(tempID)
                        # Now add the Search Root Node and the Search Name to the front of the Node List and the
                        # Clip Name to the back of the Node List
                        nodeList = (_('Search'), searchName) + nodeList + (tempID, )

                        # Add the Node to the Tree
                        self.dbTree.add_Node('SearchSnapshotNode', nodeList, line['SnapshotNum'], line['CollectNum'], sortOrder=line['SortOrder'], textSearchItems = textSearchItems)
                        # Add the Snapshot to the list of Snapshots added to the Search Result
                        addedSnapshots.append(line['SnapshotNum'])
                        
                        tmpNode = self.dbTree.select_Node(nodeList[:-1], 'SearchCollectionNode', ensureVisible=False)
                        self.dbTree.SortChildren(tmpNode)
                    # Adjust query for sqlite if needed
                    snapshotCodingQuery = DBInterface.FixQuery(snapshotCodingQuery)
                    # Execute the Snapshot Coding query
                    dbCursor.execute(snapshotCodingQuery, params)

                    # Process all results of the Snapshot Coding query 
                    for line in DBInterface.fetchall_named(dbCursor):
                        # If the Snapshot is NOT already in the Search Results ...
                        if not (line['SnapshotNum'] in addedSnapshots):
                            # Add the new Snapshot to the Database Tree Tab.
                            # To add a Snapshot, we need to build the node list for the tree's add_Node method to climb.
                            # We need to add all of the Collection Parents to our Node List, so we'll start by loading
                            # the current Collection
                            tempCollection = Collection.Collection(line['CollectNum'])

                            # Add the current Collection Node Data
                            nodeList = tempCollection.GetNodeData()                        
                            # Get the DB Values
                            tempID = line['SnapshotID']
                            # If we're in Unicode mode, format the strings appropriately
                            if 'unicode' in wx.PlatformInfo:
                                tempID = DBInterface.ProcessDBDataForUTF8Encoding(tempID)
                            # Now add the Search Root Node and the Search Name to the front of the Node List and the
                            # Clip Name to the back of the Node List
                            nodeList = (_('Search'), searchName) + nodeList + (tempID, )

                            # Add the Node to the Tree
                            self.dbTree.add_Node('SearchSnapshotNode', nodeList, line['SnapshotNum'], line['CollectNum'], sortOrder=line['SortOrder'], textSearchItems = textSearchItems)
                            # Add the Snapshot to the list of Snapshots added to the Search Result
                            addedSnapshots.append(line['SnapshotNum'])
                            
                            tmpNode = self.dbTree.select_Node(nodeList[:-1], 'SearchCollectionNode', ensureVisible=False)
                            self.dbTree.SortChildren(tmpNode)

                # If we opened a Popup Dialog, we need to close it!
                if len(textSearchItems) > 0:
                    progressDialog.Close()
                    progressDialog.Destroy()
                    
            else:
                self.searchCount = searchCount

        # If the Search Dialog is cancelled, do NOT increment the Search Number                
        else:
            self.searchCount = searchCount


    def GetSearchCount(self):
        """ This method is called to determine whether the Search Counter was incremented, that is, whether the
            search was performed or cancelled. """
        return self.searchCount


    def BuildQueries(self, queryText):
        """ Convert natural language search terms (as structured by the Transana Search Dialog) into
            executable SQL that runs on MySQL. """

        # Here are a couple of sample SQL Statements generated by this code:
        #
        # Query:  "Demo:Geometry AND NOT Demo:Teacher Commentary"
        #
        # SELECT Ep.SeriesNum, SeriesID, Ep.EpisodeNum, EpisodeID,
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Demo') AND (CK1.Keyword = 'Geometry')) THEN 1 ELSE NULL END) V1,
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Demo') AND (CK1.Keyword = 'Teacher Commentary')) THEN 1 ELSE NULL END) V2
        #   FROM ClipKeywords2 CK1, Series2 Se, Episodes2 Ep
        #   WHERE (Ep.EpisodeNum = CK1.EpisodeNum) AND (Ep.SeriesNum = Se.SeriesNum) AND (CK1.EpisodeNum > 0)
        #   GROUP BY SeriesNum, SeriesID, EpisodeNum, EpisodeID
        #   HAVING (V1 > 0) AND (V2 = 0)
        #
        # SELECT Cl.CollectNum, ParentCollectNum, Cl.ClipNum, CollectID, ClipID,
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Demo') AND (CK1.Keyword = 'Geometry')) THEN 1 ELSE NULL END) V1,
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Demo') AND (CK1.Keyword = 'Teacher Commentary')) THEN 1 ELSE NULL END) V2
        #   FROM ClipKeywords2 CK1, Collections2 Co, Clips2 Cl
        #   WHERE (Cl.ClipNum = CK1.ClipNum) AND (Cl.CollectNum = Co.CollectNum) AND (CK1.ClipNum > 0)
        #   GROUP BY Cl.CollectNum, CollectID, ClipID
        #   HAVING (V1 > 0) AND (V2 = 0)

        # Here's a query that combines Text Search and Keyword Search!
        #
        # SELECT Doc.LibraryNum, SeriesID, Doc.DocumentNum, DocumentID, 
        #        COUNT(CASE WHEN ((CK1.KeywordGroup = 'Coca Cola') AND (CK1.Keyword = 'Coke')) THEN 1 ELSE NULL END) V1, 
        #        COUNT(CASE WHEN (PlainText LIKE '%cola%') THEN 1 ELSE NULL END) V2 
        #   FROM ClipKeywords2 CK1, Series2 Se, Documents2 Doc 
        #   WHERE (Doc.DocumentNum = CK1.DocumentNum) AND 
        #         (Doc.LibraryNum = Se.SeriesNum) AND 
        #         (CK1.DocumentNum > 0) 
        #   GROUP BY Doc.LibraryNum, SeriesID, Doc.DocumentNum, DocumentID 
        #   HAVING (V1 > 0) AND (V2 > 0) 
        #   ORDER BY SeriesID, DocumentID
        
        # Initialize a Temporary Variable Counter
        tempVarNum = 0
        # We need to know if the query includes Keywords, as this alters the SQL.  This tracks that.
        includesKeywords = False
        # We also need to know if the query includes Text, as we can't do that for Snapshots.  This tracks that.
        includesText = False
        # We also need to track whether the search contains an OR operator
        includesOrOperator = False
        # We also need to keep track of what the Search Text terms are
        textSearchItems = []
        # Initialize a list for strings to store SQL "COUNT" lines
        countStrings = []
        # Initialize a list to hold the Search Parameters.
        # NOTE:  Parameters are passed separately rather than being integrated into the SQL so that
        #        MySQLdb can handle all parsing related to apostrophes and other non-SQL-friendly characters.
        params = []
        # Initialize a String to store the SQL "HAVING" clause
        havingStr = ''

        # We now will go through the Search Terms line by line and prepare to convert the Search Request to SQL
        for lineNum in range(len(queryText)):
            # Capture the Line being processed, and remove whitespace from either end
            tempStr = string.strip(queryText[lineNum])

            # Initialize the "Continuation" string, which holds a BOOLEAN Operator ("AND" or "OR")
            continStr = ''
            # Initialize the flag that signals the BOOLEAN "NOT" Operator
            notFlag = False
            # Initialize the counter that tracks the number of parentheses that are open and need to be closed.
            closeParen = 0

            # If a line ends with " AND"...
            if tempStr[-4:] == ' AND':
                # ... put the Boolean Operator into the Continuation String ...
                continStr = ' AND '
                # ... and remove it from the line being processed.
                tempStr = tempStr[:-4]

            # If a line ends with " OR"...
            if tempStr[-3:] == ' OR':
                # Note that we use an OR operator
                includesOrOperator = True
                # ... put the Boolean Operator into the Continuation String ...
                continStr = ' OR '
                # ... and remove it from the line being processed.
                tempStr = tempStr[:-3]

            # Process characters at the beginning of the Line, including open parens and the "NOT" operator.
            # NOTE:  The Search Dialog allows "(NOT", but not "NOT(".
            while (tempStr[0] == '(') or (tempStr[:4] == 'NOT '):
                # If the line starts with an open paren ...
                if tempStr[0] == '(':
                    # ... add it to the "HAVING" clause string ...
                    havingStr += '('
                    # ... and remove it from the line.
                    tempStr = tempStr[1:]
                # If the line starts with a "NOT" operator ...
                if tempStr[:4] == 'NOT ':
                    # ... set the NOT Flag ...
                    notFlag = True
                    # ... and remove it from the line.
                    tempStr = tempStr[4:]

            # Check for close parens in the line ...
            while tempStr.find(')') > -1:
                # ... keep track of how many are found in this line ...
                closeParen += 1
                # ... and remove them from the line.
                tempStr = tempStr[:tempStr.find(')')] + tempStr[tempStr.find(')') + 1:]

            # All that should be left in the line being processed now should be Keywords.
            if len(tempStr) > 0:
                # increment the Temporary Variable Counter.  (Every Keyword Group : Keyword combination gets a unique
                # Temporary Variable Number.)
                tempVarNum += 1

                # See if we have a Text Search string, either from the Search Form or the Word Frequency Report ...
                if tempStr[:20] in ['Item Text contains "', 'Word Text contains "']:
                    # Note that we are including text
                    includesText = True
                    # Remember the Text Search Term
                    textSearchItems.append(tempStr[20:tempStr.rfind('"')])
                    # Converting the Text Search Request into platform-appropriate SQL.
                    tempStr2 = "COUNT(CASE WHEN ("
                    # If we are working from Text Search from the Search Dialog ...
                    if tempStr[:20] == 'Item Text contains "':
                        # Remove the "Item Text Contains" text and the quotation marks around the search text
                        tempStr = '%%' + tempStr[20:tempStr.rfind('"')] + '%%'
                        # Find any matching text 
                        tempStr2 += "PlainText LIKE %s"
                    # If we're working from a Word Frequency Text Sarch request ...
                    else:
                        # If we're on MySQL ...
                        if TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
                            # Remove the "Item Text Contains" text and the quotation marks around the search text
                            # and add the Regular Expression code that gets whole words, even around punctuation
                            tempStr = u'([[:blank:][:punct:]]|^)' + tempStr[20:tempStr.rfind('"')] + u'([[:blank:][:punct:]]|$)'
                            # This theoretically gives whole words only  -- REGEXP '[[:<:]]%s[[:>:]]' also an option
                            tempStr2 += "PlainText REGEXP %s"
                        # if we're using SQLite ...
                        else:
                            # Remove the "Item Text Contains" text and the quotation marks around the search text
                            tempStr = '%%' + tempStr[20:tempStr.rfind('"')] + '%%'
                            # Find any matching text.  The " " || adds whole-word-only functionality to SQLite.
                            tempStr2 += '(" " || PlainText || " ") LIKE %s'
                            
                    # If we're on MySQL ...
                    if TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']:
                        # ... make the Text Search Case Insensitive!
                        tempStr2 += " COLLATE utf8_general_ci"
                    # If we're on SQLite ...
                    else:
                        # ... make the Text Search Case Insensitive!
                        tempStr2 += " COLLATE NOCASE"
                    tempStr2 += ") THEN 1 ELSE NULL END) " + "V%s" % tempVarNum
                    params.append(tempStr)

                    countStrings.append(tempStr2)
                # If not, we have KEYWORDS
                else:
                    # note that we are including Keywords
                    includesKeywords = True

                    # The presence of a variable (or it's absence if NOT has been specified) is signalled in SQL by a combination of
                    # this "COUNT" statement, which creates a numbered variable in the SELECT Clause, and a "HAVING" line.
                    # I can't adequately explain it, but it DOES work.
                    # Please, don't mess with it.

                    # Add a line to the SQL "COUNT" statements to indicate the presence or absence of a Keyword Group : Keyword pair
                    tempStr2 = "COUNT(CASE WHEN ((CK1.KeywordGroup = %s) AND (CK1.Keyword = %s)) THEN 1 ELSE NULL END) " + "V%s" % tempVarNum

                    countStrings.append(tempStr2)
                    # Add the Keyword Group to the Parameters
                    kwg = tempStr[:tempStr.find(':')]
                    if 'unicode' in wx.PlatformInfo:
                        kwg = kwg.encode(TransanaGlobal.encoding)
                    params.append(kwg)
                    # Add the Keyword to the Parameters
                    kw = tempStr[tempStr.find(':') + 1:]
                    if 'unicode' in wx.PlatformInfo:
                        kw = kw.encode(TransanaGlobal.encoding)
                    params.append(kw)
                    # Add the Temporary Variable Number that corresponds to this Keyword Group : Keyword pair to the Parameters
                    # params.append(tempVarNum)

                # If the "NOT" operator has been specified, we want the Temporary Variable to equal Zero in the "HAVING" clause
                if notFlag:
                    havingStr += '(V%s = 0)' % tempVarNum
                # If the "NOT" operator has not been specified, we want the Temporary Variable to be greater than Zero in the "HAVING" clause
                else:
                    havingStr += '(V%s > 0)' % tempVarNum

                # Add any closing parentheses that were specified to the end of the "HAVING" clause
                for x in range(closeParen):
                    havingStr += ')'
                # Add the appropriate Boolean Operator to the end of the "HAVING" clause, if one was specified
                havingStr += continStr

        # If we have Keywords AND Text Search AND an OR Operator, only items WITH SOME KEYWORDS will be found.
        # Items that contain the text but NO KEYWORDS will NOT be included in the Search results.
        # We must let the user know.
        if includesKeywords and includesText and includesOrOperator:
            msg = _('When a Search Specification includes both Keywords and Text Search \nseparated by an "OR" operator, the Search Results will not include items \nthat contain the specified text but have NO Keywords AT ALL.')
            tmpDlg = Dialogs.InfoDialog(None, msg)
            tmpDlg.ShowModal()
            tmpDlg.Destroy()

        # Before we continue, let's build the part of the query that implements the Document, Transcript, and
        # Collections selections

        # If there is a Document list, build the scoping SQL to limit which Documents are displayed
        if len(self.documentList) > 0:
            docSQL = ' AND ('
            for doc in self.documentList:
                docSQL += "(Doc.DocumentNum = %d) " % doc[0]
                if doc != self.documentList[-1]:
                    docSQL += "or "
            docSQL += ") "

        # If there is a Transcript list, build the scoping SQL to limit which Transcripts are displayed
        if len(self.transcriptList) > 0:
            transSQL = ' AND ('
            for transcript in self.transcriptList:
                transSQL += "(Tr.TranscriptNum = %d) " % transcript[0]
                if transcript != self.transcriptList[-1]:
                    transSQL += "or "
            transSQL += ") "

        # If there is a Collection list, build the scoping SQL to limit which Quotes, Clips, and Snapshots are displayed
        if len(self.collectionList) > 0:
            paramsQ = ()
            paramsCl = ()
            paramsSn = ()
            collectionSQL = ' AND ('
            for coll in self.collectionList:
                collectionSQL += "(%%s.CollectNum = %d) " % coll[0]
                if coll != self.collectionList[-1]:
                    collectionSQL += "or "
                paramsQ += ('Q',)
                paramsCl += ('Cl',)
                paramsSn += ('Sn',)
            collectionSQL += ") "

        # Now that all the pieces (countStrings, params, and the havingStr) are assembled, we can build the
        # SQL Statements for the searches.

        # Define the start of the Library/Document Query
        documentSQL = 'SELECT Doc.LibraryNum, SeriesID, Doc.DocumentNum, DocumentID, '
        # Define the start of the Library/Episode Query
        episodeSQL = 'SELECT Ep.SeriesNum, SeriesID, Ep.EpisodeNum, EpisodeID, '
        if includesText:
            episodeSQL += 'Tr.TranscriptNum, TranscriptID, '
        # Define the start of the Collection/Quote Query
        quoteSQL = 'SELECT Q.CollectNum, ParentCollectNum, Q.QuoteNum, CollectID, QuoteID, Q.SortOrder, '
        # Define the start of the Collection/Clip Query
        clipSQL = 'SELECT Cl.CollectNum, ParentCollectNum, Cl.ClipNum, CollectID, ClipID, Cl.SortOrder, '
        # Define the start of the Whole Snapshot Query
        wholeSnapshotSQL = 'SELECT Sn.CollectNum, ParentCollectNum, Sn.SnapshotNum, CollectID, SnapshotID, Sn.SortOrder, '
        # Define the start of the Snapshot Coding Query
        snapshotCodingSQL = 'SELECT Sn.CollectNum, ParentCollectNum, Sn.SnapshotNum, CollectID, SnapshotID, Sn.SortOrder, '

        # Add in the SQL "COUNT" variables that signal the presence or absence of Keyword Group : Keyword pairs or
        # text search parameters
        for lineNum in range(len(countStrings)):
            # All SQL "COUNT" lines but he last one need to end with a comma
            if lineNum < len(countStrings)-1:
                tempStr = ', '
            # The last SQL "COUNT" line does not need to end with a comma
            else:
                tempStr = ' '

            # Add the SQL "COUNT" Line and seperator to the Library/Document Query
            documentSQL += countStrings[lineNum] + tempStr
            # Add the SQL "COUNT" Line and seperator to the Library/Episode Query
            episodeSQL += countStrings[lineNum] + tempStr
            # Add the SQL "COUNT" Line and seperator to the Collection/Quote Query
            quoteSQL += countStrings[lineNum] + tempStr
            # Add the SQL "COUNT" Line and seperator to the Collection/Clip Query
            clipSQL += countStrings[lineNum] + tempStr
            if not includesText:
                # Add the SQL "COUNT" Line and seperator to the Whole Snapshot Query
                wholeSnapshotSQL += countStrings[lineNum] + tempStr
                # Add the SQL "COUNT" Line and seperator to the Snapshot Coding Query
                snapshotCodingSQL += countStrings[lineNum] + tempStr

        # Now add the rest of the SQL for the Library/Document Query
        documentSQL += 'FROM '
        if includesKeywords:
            documentSQL += 'ClipKeywords2 CK1, '
        documentSQL += 'Series2 Se, Documents2 Doc '
        documentSQL += 'WHERE '
        if includesKeywords:
            documentSQL += '(Doc.DocumentNum = CK1.DocumentNum) AND '
        documentSQL += '(Doc.LibraryNum = Se.SeriesNum) '
        if includesKeywords:
            documentSQL += 'AND (CK1.DocumentNum > 0) '
        # If there is a Document List (from Word Frequency Text Search) ..
        if len(self.documentList) > 0:
            # ... add the appropriate scoping SQL
            documentSQL += docSQL
        documentSQL += 'GROUP BY Doc.LibraryNum, SeriesID, Doc.DocumentNum, DocumentID '
        # Add in the SQL "HAVING" Clause that was constructed above
        documentSQL += 'HAVING %s ' % havingStr
        documentSQL += 'ORDER BY SeriesID, DocumentID'

        # Now add the rest of the SQL for the Library/Episode Query
        episodeSQL += 'FROM '
        if includesKeywords:
            episodeSQL += 'ClipKeywords2 CK1, '
        episodeSQL += 'Series2 Se, Episodes2 Ep'
        if includesText:
            episodeSQL += ', Transcripts2 Tr'
        episodeSQL += ' WHERE '
        if includesKeywords:
            episodeSQL += '(Ep.EpisodeNum = CK1.EpisodeNum) AND '
        episodeSQL += '(Ep.SeriesNum = Se.SeriesNum) '
        if includesKeywords:
            episodeSQL += 'AND (CK1.EpisodeNum > 0) '
        if includesText:
            episodeSQL += 'AND (Tr.EpisodeNum = Ep.EpisodeNum) AND (Tr.ClipNum = 0) '
        # If there is a Transcript List (from Word Frequency Text Search) ..
        if len(self.transcriptList) > 0:
            # ... add the appropriate scoping SQL
            episodeSQL += transSQL
        episodeSQL += 'GROUP BY Ep.SeriesNum, SeriesID, Ep.EpisodeNum, EpisodeID'
        if includesText:
            episodeSQL += ', TranscriptID'
        # Add in the SQL "HAVING" Clause that was constructed above
        episodeSQL += ' HAVING %s ' % havingStr

        # Now add the rest of the SQL for the Collection/Quote Query
        quoteSQL += 'FROM '
        if includesKeywords:
            quoteSQL += 'ClipKeywords2 CK1, '
        quoteSQL += 'Collections2 Co, Quotes2 Q '
        quoteSQL += 'WHERE '
        if includesKeywords:
            quoteSQL += '(Q.QuoteNum = CK1.QuoteNum) AND '
        quoteSQL += '(Q.CollectNum = Co.CollectNum) '
        if includesKeywords:
            quoteSQL += 'AND (CK1.QuoteNum > 0) '
        if len(self.collectionList) > 0:
            quoteSQL += collectionSQL % paramsQ
        quoteSQL += 'GROUP BY Q.CollectNum, CollectID, QuoteID '
        # Add in the SQL "HAVING" Clause that was constructed above
        quoteSQL += 'HAVING %s ' % havingStr
        # Add an "ORDER BY" Clause to preserve Quote Sort Order
        quoteSQL += 'ORDER BY CollectID, Q.SortOrder'

        # Now add the rest of the SQL for the Collection/Clip Query
        clipSQL += 'FROM '
        if includesKeywords:
            clipSQL += 'ClipKeywords2 CK1, '
        clipSQL += 'Collections2 Co, Clips2 Cl'
        if includesText:
            clipSQL += ', Transcripts2 Tr'
        clipSQL += ' WHERE '
        if includesKeywords:
            clipSQL += '(Cl.ClipNum = CK1.ClipNum) AND '
        clipSQL += '(Cl.CollectNum = Co.CollectNum) '
        if includesKeywords:
            clipSQL += 'AND (CK1.ClipNum > 0) '
        if includesText:
            clipSQL += 'AND (Tr.ClipNum = Cl.ClipNum) '
        if len(self.collectionList) > 0:
            clipSQL += collectionSQL % paramsCl
        clipSQL += 'GROUP BY Cl.CollectNum, CollectID, ClipID '
        # Add in the SQL "HAVING" Clause that was constructed above
        clipSQL += 'HAVING %s ' % havingStr
        # Add an "ORDER BY" Clause to preserve Clip Sort Order
        clipSQL += 'ORDER BY CollectID, Cl.SortOrder'

        # We can't do Snapshot searches with text!!
        if not includesText:
            # Now add the rest of the SQL for the Whole Snapshot Query
            wholeSnapshotSQL += 'FROM ClipKeywords2 CK1, Collections2 Co, Snapshots2 Sn '
            wholeSnapshotSQL += 'WHERE (Sn.SnapshotNum = CK1.SnapshotNum) AND '
            wholeSnapshotSQL += '(Sn.CollectNum = Co.CollectNum) AND '
            wholeSnapshotSQL += '(CK1.SnapshotNum > 0) '
            if len(self.collectionList) > 0:
                wholeSnapshotSQL += collectionSQL % paramsSn
            wholeSnapshotSQL += 'GROUP BY Sn.CollectNum, CollectID, SnapshotID '
            # Add in the SQL "HAVING" Clause that was constructed above
            wholeSnapshotSQL += 'HAVING %s ' % havingStr
            # Add an "ORDER BY" Clause to preserve Snapshot Sort Order
            wholeSnapshotSQL += 'ORDER BY CollectID, Sn.SortOrder'

            # Now add the rest of the SQL for the Snapshot Coding Query
            snapshotCodingSQL += 'FROM SnapshotKeywords2 CK1, Collections2 Co, Snapshots2 Sn '
            snapshotCodingSQL += 'WHERE (Sn.SnapshotNum = CK1.SnapshotNum) AND '
            snapshotCodingSQL += '(Sn.CollectNum = Co.CollectNum) AND '
            snapshotCodingSQL += '(CK1.SnapshotNum > 0) '
            # For Snapshot Coding, we ONLY want VISIBLE Keywords
            snapshotCodingSQL += 'AND (CK1.Visible = 1) '
            if len(self.collectionList) > 0:
                snapshotCodingSQL += collectionSQL % paramsSn
            snapshotCodingSQL += 'GROUP BY Sn.CollectNum, CollectID, SnapshotID '
            # Add in the SQL "HAVING" Clause that was constructed above
            snapshotCodingSQL += 'HAVING %s ' % havingStr
            # Add an "ORDER BY" Clause to preserve Snapshot Sort Order
            snapshotCodingSQL += 'ORDER BY CollectID, Sn.SortOrder'

        tempParams = ()
        for p in params:
            tempParams = tempParams + (p,)
            
##        dlg = wx.TextEntryDialog(None, "Transana Library/Document SQL Statement:", "Transana", documentSQL % tempParams, style=wx.OK)
##        dlg.ShowModal()
##        dlg.Destroy()

##        dlg = wx.TextEntryDialog(None, "Transana Library/Episode SQL Statement:", "Transana", episodeSQL % tempParams, style=wx.OK)
##        dlg.ShowModal()
##        dlg.Destroy()

##        dlg = wx.TextEntryDialog(None, "Transana Collection/Quote SQL Statement:", "Transana", quoteSQL % tempParams, style=wx.OK)
##        dlg.ShowModal()
##        dlg.Destroy()

##        dlg = wx.TextEntryDialog(None, "Transana Collection/Clip SQL Statement:", "Transana", clipSQL % tempParams, style=wx.OK)
##        dlg.ShowModal()
##        dlg.Destroy()

##        if not includesText:
##            dlg = wx.TextEntryDialog(None, "Transana Whole Snapshot SQL Statement:", "Transana", wholeSnapshotSQL % tempParams, style=wx.OK)
##            dlg.ShowModal()
##            dlg.Destroy()

##            dlg = wx.TextEntryDialog(None, "Transana Snapshot Coding SQL Statement:", "Transana", snapshotCodingSQL % tempParams, style=wx.OK)
##            dlg.ShowModal()
##            dlg.Destroy()

        # Return the Library/Episode Query, the Collection/Clip Query, the Whole Snapshot Query, the Snapshot Coding Query, 
        # and the list of parameters to use with these queries to the calling routine.
        return (documentSQL, episodeSQL, quoteSQL, clipSQL, wholeSnapshotSQL, snapshotCodingSQL, params, textSearchItems)

    def GetNodeList(self, dataTree, dataNode, nodeType):
        """ Recursively builds a list of all nodes for the Word Frequency Text Search searchScope Node
            and appropriate child nodes which match nodeType """
        # Initialize a list of results
        results = []
        # Get the First Child record
        (childNode, cookieItem) = dataTree.GetFirstChild(dataNode)
        # While there are valid Child records ...
        while childNode.IsOk():
            # ... get the node data out of the PyData
            nodeData = dataTree.GetPyData(childNode)
            # If the node is the correct type ...
            if nodeData.nodetype == nodeType:
                # ... add the node Number and node Name to the Results List
                results.append((nodeData.recNum, dataTree.GetItemText(childNode)))
            # If the Node has children ...
            if dataTree.ItemHasChildren(childNode):
                # ... recursively call this method to get the results of this node's child nodes, adding those results to these
                results += self.GetNodeList(dataTree, childNode, nodeType)
            # If this node is not the LAST child ...
            if childNode != dataTree.GetLastChild(dataNode):
                # ... then get the next child
                (childNode, cookieItem) = dataTree.GetNextChild(dataNode, cookieItem)
            # if we're at the last child ...
            else:
                # ... we can quit
                break
        # Return the results to the calling method
        return results


