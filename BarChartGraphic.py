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

"""This module creates bar chart graphics and places them in the Clipboard for use in Reports. """

__author__ = 'David K. Woods <dwoods@transana.com>'

# Import wxPython
import wx
# If we're running in stand-alone mode ...
if __name__ == '__main__':
    # ... import wxPython's RichTextCtrl
    import wx.richtext as richtext
else:
    # import Transana's Database Interface
    import DBInterface

# Import MatPlotLib's wxAgg infrastructure
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
# Import mapPlotLib's PyPlot functionality
import matplotlib.pyplot as plt


class BarChartGraphic(wx.Panel):
    """ This module accepts data, creates a BarChart from it, and places that BarChart in the CLipboard. """
    def __init__(self, parent, pos=(10, 10)):
        """ Initialize a panel for drawing the BarChart.  This panel can be hidden. """
        # Initialize a Panel
        wx.Panel.__init__(self, parent, pos=pos)
        # Create a PyPlot figure
        self.figure = plt.figure(figsize=(7, 9))
        # Create a MatPlotLib FigureCanvas based on the Panel and the Figure
        self.canvas = FigureCanvas(self, -1, self.figure)
     
    def plot(self, title, data, dataLabels):
        """ Create a BarChart.
               title        Title for the BarChart
               data         List of data values for the BarChart
               dataLabels   Matching list of data labels for the BarChart
            This module limits the BarChart to 15 bars max. """
        # Clear the Figure (in case we use the same BarChartGraphic to create multiple BarCharts)
        self.figure.clf()
        # The length of data is the number of bars we need
        numBars = len(data)
        # Define the colors to be used in the BarChart
        # If we're running stand-alone ...
        if __name__ == '__main__':
            # ... use generic colors
            colors = ['#FF0000', '#00FF00', '#0000FF', '#666600', '#FF00FF', '#00FFFF', '#440000', '#004400', '#000044']
#            colors = [(1, 0, 0), (0, 1, 0), (0, 0, 1), (0.5, 0.5, 0), (1, 0, 1), (0, 1, 1)]
            # There are no Colors for Keywords when running stand-alone
            colorsForKeywords = {}
        # If we're running inside Transana ...
        else:
            # ... import Transana's Global Values ...
            import TransanaGlobal
            # ... and use Transana's defined color scheme
            colorList = TransanaGlobal.getColorDefs(TransanaGlobal.configData.colorConfigFilename)[:-1]
            # Initialize the Colors list
            colors = []
            # Populate the colors list.
            for colorName, colorDef in colorList:
                # MatPlotLib uses a 0 .. 1 scale rather than a 0 .. 255 scale for RGB colors!
                colors.append((colorDef[0]/255.0, colorDef[1]/255.0, colorDef[2]/255.0))
            # Get the color definitions for all keywords in Transana
            colorsForKeywords = DBInterface.dict_of_keyword_colors()

        # If we have more data points that will fit, we should truncate the number of bars
        maxBars = 30
        if numBars > maxBars:
            # Reduce data to the first 15 points
            data = data[:maxBars]
            # Reduce the data labels to the first 15 labels
            dataLabels = dataLabels[:maxBars]
            # Reduce the number of bars to be displayed to maxBars
            numBars = maxBars
        # X values for the bars are simple integers for bar number
        xValues = range(numBars)
        # Set the bar width to allow space between bars
        width = .85 

        # Create a MatPlotLib SubPlot
        ax = self.figure.add_subplot(111)
        # Define the BarChart Bars in the figure
        rects1 = ax.bar(xValues, data, width)

        # Add the Chart Title
        ax.set_title(title)
        # If we're running stand-alone ...
        if __name__ == '__main__':
            # Add the Y axis label
            ax.set_ylabel('Frequency')
        else:
            # Add the Y axis label
            ax.set_ylabel(_('Frequency'))
        # Set X axis tick marks for each bar
        ax.set_xticks(xValues)
        # Add the bar labels
        lbls = ax.set_xticklabels(dataLabels, rotation=90)  # 35  65

        # Initialize the color list position indicator
        colorIndx = 0
        # For each bar ...
        for x in range(numBars):
            # If there's a color defined for the Keyword ...
            if dataLabels[x] in colorsForKeywords.keys():
                # ... use that color
                color = colorsForKeywords[dataLabels[x]]['colorDef']
            # If there's no color defined ...
            else:
                # ... use the next color from the color list
                color = colors[colorIndx]
                # Increment or reset the color index
                if colorIndx >= len(colors) - 1:
                    colorIndx = 0
                else:
                    colorIndx += 1
            # ... define the bar color
            rects1[x].set_color(color)
            # ... make the label color match the bar color
            lbls[x].set_color(color)
        # Give the graph small inside margins
        plt.margins(0.05)
        # Adjust the bottom margin to make room for bar labels
        plt.subplots_adjust(bottom=0.5)  # 0.2  0.4
        # Draw the BarChart
        self.canvas.draw()
        # Copy the BarChart to the Clipboard
        self.canvas.Copy_to_Clipboard()


if __name__ == '__main__':
    class TestFrame(wx.Frame):
        def __init__(self,parent,title):
            wx.Frame.__init__(self,parent,title=title,size=(1300,1000))
             
            # Placed up front so other screen elementes will be placed OVER it!
            self.barChartGraphic = BarChartGraphic(self)  # DON'T PUT THIS IN THE SIZER -- It needs to be hidden!

            # Create a Sizer
            s1 = wx.BoxSizer(wx.HORIZONTAL)
            # Add a RichTextCtrl
            self.txt = richtext.RichTextCtrl(self, -1)
            # Put the RichTextCtrl on the Sizer
            s1.Add(self.txt, 1, wx.EXPAND)
            # Set the main sizer
            self.SetSizer(s1)
            # Lay out the window
            self.Layout()
            self.SetAutoLayout(True)

            self.txt.AppendText('Now comes the fun part.\n\n')

            # Prepare some fake data for a demonstration graph
            title = 'Keyword Frequency'
            data = [50, 45, 25, 22, 20, 18, 15, 12, 11, 8, 7, 6, 4, 2, 1, 1]
            dataLabels = ['Long Label 1', 'Long Label 2', 'Really, really long label 3', 'Long Label 4', 'Long Label 5', 'Long Label 6', 'Long Label 7',
                          'Long Label 8', 'Long Label 9', 'Long Label 10', 'Long Label 11', 'Long Label 12', 'Long Label 13', 'Long Label 14',
                          'Long Label 15', '16']

            # Draw the BarChart.  This places the graphic in the Clipboard.
            self.barChartGraphic.plot(title, data, dataLabels)

            # If the Clipboard isn't Open ...
            if not wx.TheClipboard.IsOpened():
                # ... open it!
                clip = wx.Clipboard()
                clip.Open()

                # Create an Image Data Object
                bitmapObject = wx.BitmapDataObject()
                # Get the Data from the Clipboard
                clip.GetData(bitmapObject)
                # Convert the BitmapsDataObject to a Bitmap
                bitmap = bitmapObject.GetBitmap()
                # Convert the Bitmap to an Image
                image = bitmap.ConvertToImage()

                # Write the plain text into the Rich Text Ctrl
                # self.txt.WriteBitmap(bitmap, wx.BITMAP_TYPE_BMP) BAD FOR RTF
                self.txt.WriteImage(image)

                self.txt.AppendText('Bitmap Added!!\n\n')

                # Close the Clipboard
                clip.Close()

            # Draw a second BarChart
            # Prepare some fake data for a demonstration graph
            title = 'Keyword Frequency'
            data = [25, 22, 20, 18, 15]
            dataLabels = ['Long Label 4', 'Long Label 5', 'Long Label 6', 'Long Label 7',
                          'Long Label 8']
            self.barChartGraphic.plot(title, data, dataLabels)

            self.barChartGraphic.canvas.Copy_to_Clipboard()

            # If the Clipboard isn't Open ...
            if not wx.TheClipboard.IsOpened():
                # ... open it!
                clip = wx.Clipboard()
                clip.Open()

                # Create an Image Data Object
                bitmapObject = wx.BitmapDataObject()
                # Get the Data from the Clipboard
                clip.GetData(bitmapObject)
                # Convert the BitmapsDataObject to a Bitmap
                bitmap = bitmapObject.GetBitmap()
                # Convert the Bitmap to an Image
                image = bitmap.ConvertToImage()

                # Write the plain text into the Rich Text Ctrl
                # self.txt.WriteBitmap(bitmap, wx.BITMAP_TYPE_BMP) BAD FOR RTF
                self.txt.WriteImage(image)

                self.txt.AppendText('Bitmap Added!!\n\n')

                # Close the Clipboard
                clip.Close()


    # Initialize the App when in stand-alone mode
    app = wx.App()
    frame = TestFrame(None, "BarChart in RichTextCtrl")
    frame.Show()
    app.MainLoop()
