# VBA-MACROS-EXCEL-

The project was based on MACROS/VBA for merging two files and Creating a random username and password for unlimited data of customers.  The project was done using VBA excel and some other tools in excel. 

The final output was just 3 buttons that need to press to do the following tasks:

1- TO copy data from User file to Merged File
2- To copy data from Organizational data to the merged file
3- A button to create a Random Username and Random password for millions of members.
==============================================================================

VBA Details: 

Sub Macro9()
'
' Macro9 Macro
'

'
    Range("Table_2[@Column5]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(LEFT(RC[-4],20),CHOOSE(RANDBETWEEN(1,35),1,2,3,4,5,6,7,8,9,""a"",""b"",""c"",""d"",""e"",""f"",""g"",""h"",""i"",""j"",""k"",""l"",""m"",""n"",""o"",""p"",""q"",""r"",""s"",""t"",""u"",""v"",""w"",""x"",""y"",""z""),CHOOSE(RANDBETWEEN(1,35),1,2,3,4,5,6,7,8,9,""a"",""b"",""c"",""d"",""e"",""f"",""g"",""h"",""i"",""j"",""k"",""l"",""m"",""n"",""o"",""p""," & _
        """q"",""r"",""s"",""t"",""u"",""v"",""w"",""x"",""y"",""z""),CHOOSE(RANDBETWEEN(1,35),1,2,3,4,5,6,7,8,9,""a"",""b"",""c"",""d"",""e"",""f"",""g"",""h"",""i"",""j"",""k"",""l"",""m"",""n"",""o"",""p"",""q"",""r"",""s"",""t"",""u"",""v"",""w"",""x"",""y"",""z""),CHOOSE(RANDBETWEEN(1,35),1,2,3,4,5,6,7,8,9,""a"",""b"",""c"",""d"",""e"",""f"",""g"",""h"",""i"",""j"",""k" & _
        """,""l"",""m"",""n"",""o"",""p"",""q"",""r"",""s"",""t"",""u"",""v"",""w"",""x"",""y"",""z""))" & _
        ""
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Range("Table_2[@Column6]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveCell.FormulaR1C1 = _
        "=CHAR(RANDBETWEEN(65,90))&CHAR(RANDBETWEEN(65,90))&CHAR(RANDBETWEEN(97,122))&CHAR(RANDBETWEEN(97,122))&CHAR(RANDBETWEEN(35,38))&RANDBETWEEN(1111,9999)"
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FillDown
    Columns("G:G").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub




Sub Macro10()
'
' Macro10 Macro
'
'
    Sheets("Adult User Information").Select
    Range("Table_1[@Column1]").Select
    ActiveCell.Offset(1, 0).Range("Table_1[@Column1]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Final").Select
    Range("Table_2[@Column1]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
End Sub




Sub Macro44()
'
' Macro44 Macro
'

'
    Sheets("Organization Information").Select
    Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Final").Select
    Range("Table_2[@Column8]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.FillDown
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    Range("I9").Select
    Sheets("Organization Information").Select
    Range("A3").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Final").Select
    Range("Table_2[@Column10]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.FillDown
    Sheets("Organization Information").Select
    Range("A5").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Final").Select
    Range("Table_2[@Column11]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.FillDown
    Range("K8").Select
    Sheets("Organization Information").Select
    Range("A7").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Final").Select
    Range("Table_2[@Column12]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.FillDown
    Sheets("Organization Information").Select
    Range("A9").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Final").Select
    Range("Table_2[@Column13]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.FillDown
    Sheets("Organization Information").Select
    Range("A11").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Copy
    Sheets("Final").Select
    Range("Table_2[@Column14]").Select
    ActiveCell.Offset(1, 0).Range("Table_2[@Column1]").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.FillDown
    Range("Table_2[#All]").Select
    Range("N2").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
End Sub
