Sub IEEE_DoubleColumn_Body_Macro()
'
'Kalkidan Zeberega
'
'This macro formats the body text of a document to match the formatting used by IEEE Double Column Conference Papers.
'
' Select the body text before running the macro. All formatting will then be automatically taken care of.

    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 10
    WordBasic.PageSetupMargins Tab:=0, PaperSize:=0, TopMargin:="0.75", _
        BottomMargin:="1", LeftMargin:="0.63", RightMargin:="0.63", Gutter:="0", _
        PageWidth:="8.5", PageHeight:="11", Orientation:=0, FirstPage:=0, _
        OtherPages:=0, VertAlign:=0, ApplyPropsTo:=4, FacingPages:=0, _
        HeaderDistance:="0.5", FooterDistance:="0.5", SectionStart:=2, _
        OddAndEvenPages:=0, DifferentFirstPage:=0, Endnotes:=0, LineNum:=0, _
        CountBy:=0, TwoOnOne:=0, GutterPosition:=0, LayoutMode:=0, DocFontName:= _
        "", FirstPageOnLeft:=0, SectionType:=1, FolioPrint:=0, ReverseFolio:=0, _
        FolioPages:=1
    With Selection.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = False
        .Width = InchesToPoints(3.5)
        .Spacing = InchesToPoints(0.24)
    End With
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
End Sub
