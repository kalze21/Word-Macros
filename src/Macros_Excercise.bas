Attribute VB_Name = "NewMacros"
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro4"
'
' Macro4 Macro
'
'
    Selection.Font.Italic = True
    ActiveDocument.Save
    ActiveDocument.Save
    ActiveDocument.Save
    ActiveDocument.Save
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro5"
'
' Macro5 Macro
'
'
    Selection.Font.Italic = True
End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Selection.Font.Bold = True
    Selection.Font.Size = 9
    Windows("2014_04_msw_usltr_format (2) [Compatibility Mode]").Activate
    Windows("Kalkidan Zeberega_Macro.").Activate
    ActiveDocument.Save
    ActiveDocument.Save
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
    Selection.Font.Bold = True
    Windows("2014_04_msw_usltr_format (2) [Compatibility Mode]").Activate
    Windows("Kalkidan Zeberega_Macro.").Activate
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro6"
'
' Macro6 Macro
'
'
    Selection.Font.Size = 10
    ActiveDocument.Save
End Sub
