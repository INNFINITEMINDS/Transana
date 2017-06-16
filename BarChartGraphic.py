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

"""This module creates a bar chart graphic and makes the bitmap available. """

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


class BarChartGraphic(object):
    """ This module accepts data, creates a BarChart from it, and returns a wx.Bitmap. """
    def __init__(self, title, data, dataLabels, size=(800, 700)):
        """ Create a BarChart.
               title        Title for the BarChart
               data         List of data values for the BarChart
               dataLabels   Matching list of data labels for the BarChart
            This module sizes the bitmap for a printed page by default. """

        def barXPos(x):
            """ Calculate the horizontal center for each bar in pixels """
            # numBars, barChartLeft, and barChartWidth are constants in the calling routine
            # Calculate the width of each bar
            barWidth = barChartWidth / numBars
            # Calculate the position of each bar
            barXPos = (float(x) / numBars) * barChartWidth + barWidth / 2 + barChartLeft
            # Return the bar center position
            return barXPos

        def barHeight(x):
            """ Calculate the height of each bar in pixels """
            # data and barChartHeight are constants in the calling routine
            # Determine the size of the largest bar
            barMax = max(data)
            # Calculate the height of the bar value passed in
            barHeight = float(x) / barMax * barChartHeight
            # We return 95% of the height value to give the bar chart some white space at the top.
            return int(barHeight * 0.95)

        def verticalAxisValues(maxVal):
            """ Given the maximum value of the axis, determine what values should appear as axis labels.
                This method implements the 2-5-10 rule. """
            # Initialize a list of values to return
            values = []

            # Let's normalize the data as part of assigning axis labels
            # Initilaize the increment between axis label values
            increment = 1
            # Initialize the conversion factor for axis labels, used in handling large values
            convertFactor = 1
            # While our maxValue is over 100 ...
            while maxVal > 100:
                # ... reduce the maximum value by a factor of 10 ...
                maxVal /= 10
                # ... and increase our conversion factor by a factor of 10.
                convertFactor *= 10
            # If our normalized max value is over 50 ...
            if maxVal > 50:
                # ... increments of 10 will give us between 5 and 10 labels
                increment = 10
            # If our normalized max value is between 20 and 50 ...
            elif maxVal > 20:
                # ... increments of 5 will give us between 4 and 10 labels
                increment = 5
            # If our normalized max value is between 8 and 20 ...
            elif maxVal > 8:
                # ... increments of 2 will give us between 4 and 10 labels
                increment = 2
            # If our normalized max value is 8 or less ...
            else:
                # ... increments of 1 will five us between 1 and 8 labels.
                increment = 1

            # for values between 0 and our max value (plus 1 to include the mac value if a multiple of 10) space by increment ...
            for x in range(0, maxVal + 1, increment):
                # ... add the incremental value multiplied by the conversion factor to our list of axis labels
                values.append(x * convertFactor)
            # Return the list of axis labels
            return values
            
        # Get the Graphic Dimensions
        (imgWidth, imgHeight) = size

        # Create an empty bitmap
        self.bitmap = wx.EmptyBitmap(imgWidth, imgHeight)
        # Get the Device Context for that bitmap
        self.dc = wx.BufferedDC(None, self.bitmap)
     
        # The length of data is the number of bars we need
        numBars = len(data)

        # Determine the longest bar label
        # Define the label font size
        axisLabelFontSize = 11
        # Define a Font for axis labels
        font = wx.Font(axisLabelFontSize, wx.FONTFAMILY_DEFAULT, wx.NORMAL, wx.NORMAL, False)
        # Set the Font for the DC
        self.dc.SetFont(font)
        # Initize the max width variable
        maxWidth = 0
        # For each bar label ...
        for x in range(numBars):
            # ... determine the size of the label
            (txtWidth, txtHeight) = self.dc.GetTextExtent(dataLabels[x])
            # See if it's bigger than previous labels
            maxWidth = max(txtWidth, maxWidth)

        # Give a left margin of 70 pixels for the vertical axis labels
        barChartLeft = 70
        # The width of the bar chart will be the image width less the left margin and 25 pixels for right margin
        barChartWidth = imgWidth - barChartLeft - 25
        # Give a top margin of 50 pixels to have room for the chart title
        if title != '':
            barChartTop = 50
        # or 20 pixels if there is no title
        else:
            barChartTop = 20
        # Reserve almost HALF the image for bar labels.  (Transana uses LONG labels!)
        barChartHeight = max(imgHeight / 2, imgHeight - maxWidth - barChartTop - 30)
        
        # Initialize a colorIndex to track what color to assign each bar
        colorIndx = 0
        # Define the colors to be used in the BarChart
        # If we're running stand-alone ...
        if __name__ == '__main__':
            # ... use generic colors
            colors = ['#FF0000', '#00FF00', '#0000FF', '#666600', '#FF00FF', '#00FFFF', '#440000', '#004400', '#000044']
            # There are no defined Colors for Keywords when running stand-alone
            colorsForKeywords = {}
        # If we're running inside Transana ...
        else:
            # ... import Transana's Global Values ...
            import TransanaGlobal
            # ... and use Transana's defined color scheme
            colorList = TransanaGlobal.getColorDefs(TransanaGlobal.configData.colorConfigFilename)[:-1]
            # Initialize the Colors list
            colors = []
            # Populate the colors list from Transana's defined colors
            for colorName, colorDef in colorList:
                colors.append(colorDef)
            # Get the color definitions for all keywords in Transana with defined colors
            colorsForKeywords = DBInterface.dict_of_keyword_colors()

        # Define a white brush as the DC Background
        self.dc.SetBackground(wx.Brush((255, 255, 255)))
        # Clear the Image (uses the Brush)
        self.dc.Clear()

        # Draw a border around the whole bitmap
        # Set the DC Pen
        self.dc.SetPen(wx.Pen((0, 0, 0), 2, wx.SOLID))
        # Draw an outline around the whole graphic
        self.dc.DrawRectangle(1, 1, imgWidth - 1, imgHeight - 1)

        # Place the Title on the DC
        # Define a Font
        font = wx.Font(18, wx.FONTFAMILY_DEFAULT, wx.NORMAL, wx.NORMAL, False)
        # Set the Font for the DC
        self.dc.SetFont(font)
        # Set the Text Foreground Color to Black
        self.dc.SetTextForeground(wx.Colour(0, 0, 0))
        # Determine the size of the title text
        (titleWidth, titleHeight) = self.dc.GetTextExtent(title)
        # Add the Title to the Memory DC (and therefore the bitmap)
        self.dc.DrawText(title, imgWidth / 2 - titleWidth / 2, 10)

        # Define a Font for axis labels
        font = wx.Font(axisLabelFontSize, wx.FONTFAMILY_DEFAULT, wx.NORMAL, wx.NORMAL, False)
        # Set the Font for the DC
        self.dc.SetFont(font)

        # Draw Axes
        # Draw an outline around the Bar Chart area with just a little extra width so it looks better
        self.dc.DrawRectangle(barChartLeft - 3, barChartTop, barChartWidth + 6, barChartHeight)
        # Get the values that should be used for the vertical axis labels
        axisValues = verticalAxisValues(max(data))
        # For each axis label ...
        for x in axisValues:
            # ... draw a pip at the value
            self.dc.DrawLine(barChartLeft - 8, barChartTop + barChartHeight - barHeight(x) - 1, barChartLeft - 3, barChartTop + barChartHeight - barHeight(x) - 1)
            # Convert the axis value to right-justified text
            axisLbl = "%10d" % x
            # Determine the size of the axis label
            (txtWidth, txtHeight) = self.dc.GetTextExtent(axisLbl)
            # Add the text to the drawing at just the right position
            self.dc.DrawText(axisLbl, barChartLeft - txtWidth - 13, barChartTop + barChartHeight - barHeight(x) - txtHeight / 2)

        # For each bar in the bar chart ...
        for x in range(numBars):
            # ... draw the pips for the bar labels
            self.dc.DrawLine(barXPos(x), barChartTop + barChartHeight, barXPos(x), barChartTop + barChartHeight + 3)
            
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

            # Label the bars
            # Set the Text Foreground Color
            self.dc.SetTextForeground(color)
            # Get the size of the bar label
            (txtWidth, txtHeight) = self.dc.GetTextExtent(dataLabels[x])
            # Add the bar label to the bar chart
            self.dc.DrawRotatedText(dataLabels[x], barXPos(x) + (txtHeight / 2), barChartTop + barChartHeight + 10, 270)
            # Calculate bar width as 80% of the width a bar would be with no whitespace between bars
            barWidth = float(barChartWidth) / numBars * 0.8
            # Set the Brush color to the color to be used for the bar
            self.dc.SetBrush(wx.Brush(color))
            # Draw the actual bar.  3 extra points compensates for the line thickness.
            self.dc.DrawRectangle(barXPos(x) - barWidth / 2 + 1,
                                  barChartTop + barChartHeight - barHeight(data[x]) - 3,
                                  barWidth,
                                  barHeight(data[x]) + 3)

    def GetBitmap(self):
        """ Provide the Bitmap to the calling routine """
        # Return the Bitmap object as applied 
        return self.bitmap


if __name__ == '__main__':
    class TestFrame(wx.Frame):
        def __init__(self,parent,title):
            wx.Frame.__init__(self,parent,title=title,size=(1300,1000))
             
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


            # Draw the BarChart.  This places the graphic in the Report.
            bc1 = BarChartGraphic(title, data, dataLabels)
            bitmap1 = bc1.GetBitmap()
            self.txt.WriteImage(bitmap1.ConvertToImage())


            # Draw a second BarChart
            # Prepare some fake data for a demonstration graph
            title = ''
            data = [25, 22, 20, 18, 15]
            dataLabels = ['Long Label 4', 'Long Label 5', 'Long Label 6', 'Long Label 7',
                          'Long Label 8']
            bc1 = BarChartGraphic(title, data, dataLabels)
            bitmap1 = bc1.GetBitmap()
            self.txt.WriteImage(bitmap1.ConvertToImage())


    # Initialize the App when in stand-alone mode
    app = wx.App()
    frame = TestFrame(None, "BarChart in RichTextCtrl")
    frame.Show()
    app.MainLoop()
