Attribute VB_Name = "Module2"
Sub copyB()
Attribute copyB.VB_ProcData.VB_Invoke_Func = " \n14"
'
' copyB Macro
'

'
    Range("AA2:AX25").Select
    Range("AX2").Activate
    Selection.Copy
    Range("B2:Y25").Select
    ActiveSheet.Paste
End Sub
