Attribute VB_Name = "NewMacros"
'Kalkidan Zeberega
'EE393

'This is a code for the body part

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Selection.Font.Name = "Times New Roman"
    WordBasic.PageSetupMargins Tab:=0, PaperSize:=0, TopMargin:="0.75", _
        BottomMargin:="1", LeftMargin:="0.63", RightMargin:="0.63", Gutter:="0", _
        PageWidth:="8.5", PageHeight:="11", Orientation:=0, FirstPage:=0, _
        OtherPages:=0, VertAlign:=0, ApplyPropsTo:=4, FacingPages:=0, _
        HeaderDistance:="0.5", FooterDistance:="0.5", SectionStart:=2, _
        OddAndEvenPages:=0, DifferentFirstPage:=0, Endnotes:=0, LineNum:=0, _
        CountBy:=0, TwoOnOne:=0, GutterPosition:=0, LayoutMode:=0, DocFontName:= _
        "", FirstPageOnLeft:=0, SectionType:=1, FolioPrint:=0, ReverseFolio:=0, _
        FolioPages:=1
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type <> wdPrintView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = False
        .Width = InchesToPoints(3.5)
        .Spacing = InchesToPoints(0.24)
    End With
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    ActiveWindow.ActivePane.VerticalPercentScrolled = 3
    ActiveWindow.ActivePane.VerticalPercentScrolled = 33
    ActiveWindow.ActivePane.VerticalPercentScrolled = 12
    ActiveWindow.ActivePane.VerticalPercentScrolled = 44
    ActiveWindow.ActivePane.VerticalPercentScrolled = 16
End Sub
