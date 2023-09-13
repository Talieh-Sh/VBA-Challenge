Attribute VB_Name = "Module1"
Sub MakeBold()
Attribute MakeBold.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MakeBold Macro
'

'
    Range("A1:A6").Select
    Selection.Font.Bold = True
End Sub
Sub FontChange()
Attribute FontChange.VB_ProcData.VB_Invoke_Func = " \n14"
'
' FontChange Macro
'

'
    With Selection.Font
        .Name = "AkayaTelivigala-Regular"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
