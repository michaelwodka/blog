---
layout: post
title: 5 ways to automate Microsoft Excel through VBA
description: Tired of doing tedious, repetitive tasks in Excel? If you're a heavy user of Excel and don't know how to use Visual Basic for Applications (VBA), you're seriously killing your productivity. Save hundreds of hours of Excel work by learning how to code in VBA.
author: Michael Wodka
permalink: /five-ways-to-automate-excel-through-vba/
imgurl1: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-image1.1.png
imgurl2: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-image2.png
imgurl3: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-image3.png
gifurl1: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif1.gif
gifurl2: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif2.gif
gifurl3: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif3.gif
gifurl4: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif4.gif
gifurl5: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif5.gif
gifurl6: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif6.gif
gifurl7: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif7.gif
gifurl8: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif8.gif
gifurl9: /assets/2018-09-06-top-five-ways-to-automate-excel/post2-gif9.gif
---

![]({{page.imgurl1|relative_url}})

Microsoft Excel. Love it or hate it - there is no denying the dominance of this spreadsheet software across businesses around the world. Millions of people use it everyday from accounting to business intelligence to financial modeling. While most people have at least a basic understanding of how to use Excel (e.g., inputting data into cells, creating graphs, making calculations), a large majority have never tapped into Excel's true potential as a workflow automation machine. 

If you don't already know, you can use Excel to automatically create PowerPoint presentations, send Outlook emails, move and copy files around your computer, and so much more. This can all be done through Visual Basic for Applications (VBA), the programming language of Excel and other Microsoft Office products.

Now for my non-technical audience, performing basic automation tasks in VBA is actually not that hard. In fact, I will be walking you through the easiest way to utilize VBA: recording a macro, which requires no coding whatsoever! If you are already familiar with VBA, I encourage you to scroll down to the use case sections to see some excellent applications of VBA.

If you're still not convinced on the merits of learning Excel VBA, all I will say is if you've ever done a tedious, repetitive task in Excel such as updating a spreadsheet with new data, using VBA will save you hundreds of hours of work in the long run and make your life much easier.

## Setting up VBA in Excel

###### _**Important Note**: this tutorial is assuming that you're working on Windows 10 with Microsoft Excel 2016. If you're using an older version of Windows or Excel, most of this tutorial should still work for you (though not guaranteed). I also can't guarantee this tutorial will work with Excel on macOS. In addition, you need to ensure macros are enabled in Excel for VBA to work in your environment._

Setting up Excel VBA is relatively simple. All you have to do is click the "Developer" tab in the ribbon, and you will gain access to writing VBA scripts and recording macros.

![]({{page.imgurl2|relative_url}})

If the "Developer" tab is not showing up in your ribbon, no need to panic! To add it, just click the "File" tab from the ribbon and then the "Options" button on the next screen in the left-hand column. In the "Options" window, click the "Customize Ribbon" option in the left-hand column and on the next screen, click the checkbox next to "Developer" under the "Main Tabs" column and click "Ok" at the bottom. The "Developer" tab should now appear in your ribbon.

![]({{page.gifurl1|relative_url}})

## Recording your first macro

A macro is an action or a set of actions that you can run as many times as you want in Excel. You can either code a macro using Excel VBA or simply record a set of actions you manually perform in Excel and assign them to a macro.

For my non-technical audience, recording a macro is the easiest way to get started with Excel VBA, and it requires no coding whatsoever. All you need to do is click the "Developer" tab in the ribbon, click the button "Use Relative References" and when you're ready to start recording, click the "Record a Macro" button in the ribbon.

You will then get a pop-up prompting you to name your macro, write a description, and call it using a shortcut. For the purposes of this tutorial, just click the "Ok" button when you see the screen and keep the default name as "macro1"

From this point forward, every action you perform in Excel will now be recorded (e.g., entering data into cells, copying data into different worksheets) and assigned to a macro.

To see how it works, write the numbers "1", "2", and "3" in cells "A1", "A2", and "A3" respectively. It's important to avoid mouse clicking as much as possible since this can potentially screw up the macro when applying it to different cells and worksheets. When performing the task above, simply click the "Enter" button on your keyboard to get to the next cell.

Once you're done, click the "Stop Recording" button and your macro will then be created! To run it, click on a new cell and then click the "Macros" button in the ribbon. In the pop-up screen, select your macro and click "Run". "1", "2", and "3" should now appear below your active cell.

![]({{page.gifurl2|relative_url}})

To make running macros even easier, I recommend you assign the macro to a shape, so when you click on the shape, the macro will run automatically. To do this, go to "Insert" in the ribbon, click the "Shapes" option, and choose the shape you want to add. Once the shape has been created, right click on it and choose the "Assign Macro" button. 

In the pop-up screen, choose your macro to assign and click the "Ok" button. Now everytime you click the shape, it will act like a button to run your macro. You can can even add text to the shape like "Run Macro" to really make it look like a button.

![]({{page.imgurl3|relative_url}})

And it's that simple. You can use the "Record a Macro" feature for a number of different use cases. I encourage you to test it out.

## Writing your first VBA script

Now if you want to get into actual coding and more dynamic use cases for VBA, you will have to learn how to write a VBA script. To do so, just click the "Developer" tab in the ribbon and then click the "Visual Basic" icon in the left-hand area of the ribbon.

The next screen will look pretty complex, but there is really only one section you need to understand to get started. In the left-hand column, you will find a folder called "Microsoft Excel Objects". Right click on this folder and then click "Insert" and then click "Module".

A blank sheet will then appear on your screen. This is where you will write your VBA code.

![]({{page.gifurl3|relative_url}})

Now we get to the fun part. Below is the code that replicates the exercise we did above when recording the macro to input "1", "2", "3" in consecutive rows.

```
Public Sub firstscript()
	ActiveCell.Value = 1
	ActiveCell.Offset(1, 0) = 2
	ActiveCell.Offset(2, 0) = 3
End Sub
```

Paste this code directly into the blank sheet and then click green arrow in the toolbar to run the script. See below for a visual of the code execution.

![]({{page.gifurl4|relative_url}})

Now you may have a lot of questions around the syntax of the script and how this script even works. Rather than get into the details in this article, I recommend you follow a more in-depth tutorial to learn how to write VBA. I highly recommend this [**excellent YouTube VBA tutorial series by WiseOwlTutorials**](https://www.youtube.com/watch?v=KHO5NIcZAc4&list=PLNIs-AWhQzckr8Dgmgb3akx_gFMnpxTN5). I used this exact tutorial to learn VBA!

If you're interested in seeing more use cases for Excel VBA, check out the five sections below.

## 1) Running a calculation for cells across multiple worksheets

Let's say you have a value in cell "A1" across five worksheets that you want to aggregate (e.g., get the sum). See below for a calculation VBA script that will aggregate these values and output the sum to cell "A1" in "Sheet6".

```
Public Sub sumOfCells()
	For x = 1 To 5
	    finalValue = Sheets(x).Range("A1").Value + finalValue
	Next x

	Sheets(6).Range("B1") = finalValue
End Sub
```

![]({{page.gifurl5|relative_url}})

## 2) Copy a worksheet from one workbook to another

Ever had to merge two workbooks together or copy a worksheet from one workbook to another? If so, see below for a VBA script that copies "Sheet2" from "Workbook2" and pastes it into "Workbook1".

```
Public Sub copyWorksheet()
   Workbooks("Workbook2").Sheets(1).Copy _
   After:=Workbooks("Workbook1").Sheets(1)
End Sub
```

![]({{page.gifurl6|relative_url}})

## 3) Create a pie chart

As you progress to more advanced Excel VBA, learning how to automatically create graphs and charts can be very useful. Check out the VBA script below to create a simple pie chart.

```
Public Sub createPieChart()
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlPie
    ActiveChart.SetSourceData Source:=Range("A1:B2")
    ActiveChart.Parent.Name = "Pie Chart"
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    ActiveChart.ChartTitle.Text = "Pie Chart"
    Selection.Format.TextFrame2.TextRange.Font.Size = 10
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Size = 8
    ActiveChart.SeriesCollection(1).DataLabels.Select
    Selection.Format.TextFrame2.TextRange.Font.Size = 8
    ActiveChart.ChartArea.Select
End Sub
```

![]({{page.gifurl7|relative_url}})

## 4) Create a PowerPoint slide

One of the best features of VBA is that you can use it across several Microsoft Office products. This can be very useful since many people use Excel to analyze data and create charts for PowerPoint presentations. Check out the VBA script below to export the pie chart we created in the previous script to a PowerPoint slide.

```
Public Sub createPPT()
    Dim newPowerPoint As Object
    Dim myPresentation As Object
    Dim mySlide As Object
    
    ActiveSheet.ChartObjects("Pie Chart").Activate
    ActiveChart.ChartArea.Copy
    
    If newPowerPoint Is Nothing Then
        Set newPowerPoint = CreateObject("PowerPoint.Application")
    End If

    If newPowerPoint.Presentations.Count = 0 Then
        Set myPresentation = newPowerPoint.Presentations.Add
    End If

    newPowerPoint.Visible = True
    
    Set mySlide = myPresentation.Slides.Add(1, 11) '11 = ppLayoutTitleOnly
    
    mySlide.Shapes.PasteSpecial DataType:=2  '2 = ppPasteEnhancedMetafile
    Set myShape = mySlide.Shapes(mySlide.Shapes.Count)
    
    myShape.Left = 66
    myShape.Top = 152
    myShape.Height = 300
    myShape.Width = 600
    
    mySlide.Shapes(1).TextFrame.TextRange.Text _
    = "Percentage of people who know VBA!"
End Sub
```

![]({{page.gifurl8|relative_url}})

## 5) Draft an Outlook email

If you liked creating PowerPoint slides through Excel VBA, you'll love using it to draft Outlook emails! Check out the VBA script below to export the pie chart we created in the previous script to an Outlook draft email.

```
Public Sub draftEmail()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim wordDoc As Object
    
    ActiveSheet.ChartObjects("Pie Chart").Activate
    ActiveChart.ChartArea.Copy
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = "test@email.com"
        .Subject = "Percentage of People Who Know VBA!"
        .Display
        
        Set wordDoc = OutMail.GetInspector.WordEditor
        wordDoc.Range.PasteAndFormat wdChartPicture
    End With
End Sub
```

![]({{page.gifurl9|relative_url}})