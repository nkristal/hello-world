Private Sub cmdFill_Click()
    checkInput
End Sub



Private Sub UserForm_Initialize()

'init dates comboboxes
    For i = 1 To 31
       cbxDay.AddItem i
    Next i
    
    For j = 1980 To 2024
       cbxYear.AddItem j
    Next j
    
    cbxMonth.List = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
    cbxPhoneCode.List = Array("02", "03", "04", "08", "050", "052", "053", "054", "058", "057", "059")
    
'init text boxes
    txtPName = ""
    txtLName = ""
    txtAddress = ""
    txtPhone = ""

'init radio buttons
    optFemale = False
    optMale = False
    
'init check boxes
    chbACC = False
    chbADV = False
    chbVBA = False
    
    
End Sub

Function checkInput()
' Verify phone number
    check = True
    txtPhone.BackColor = RGB(196, 255, 196)
    cbxPhoneCode.BackColor = RGB(196, 255, 196)
    frmCourses.BackColor = RGB(196, 255, 196)
    
    If cbxPhoneCode = "" Then
        check = False
        cbxPhoneCode.BackColor = RGB(255, 128, 128)
        Debug.Print "no phone code"
    End If
    
    If Not IsNumeric(txtPhone) Then
        check = False
        txtPhone.BackColor = RGB(255, 128, 128)
        Debug.Print "not numeric"
        Else
            If (txtPhone < 1000000) Or (txtPhone > 9999999) Then
                check = False
                txtPhone.BackColor = RGB(255, 128, 128)
                Debug.Print "not a phone number"
            End If
        
    End If
    
    
' Verify at least one course is selected
    If Not (chbACC Or chbADV Or chbVBA) Then
        Debug.Print "no course selected"
        frmCourses.BackColor = RGB(255, 128, 128)
    End If
    
    

' return TRUE if input is correct
    checkInput = check
End Function
Sub fillTable()
' äæðú ùåøä øé÷ä
    If Cells(2, 1) <> "" Then
        Rows("2:2").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    End If
    
'äæðú îöééï ùåøä øõ
    Cells(2, 1) = 1 + Cells(3, 1)
    
'äæðú äðúåðéí ìèáìä
    Cells(2, 2) = Date
    Cells(2, 3) = Time
    
    Cells(2, 4) = txtPName
    Cells(2, 5) = txtLName
    Cells(2, 6) = txtAddress
    Cells(2, 7) = txtPhone
    
    If optMale Then Cells(2, 8) = "æëø"
    If optFemale Then Cells(2, 8) = "ð÷áä"
    
    Cells(2, 10) = chbVBA
    Cells(2, 11) = chbADV
    Cells(2, 12) = chbACC
    
    
End Sub
Private Sub cmdFillExit_Click()
    'checkInput
    fillTable

'Closing the form
    Debug.Print chbACC
    Debug.Print chbADV
    Debug.Print chbVBA
    
'    MsgBox "úåãä òì ôðééúê!"
    Unload Me
'    ThisWorkbook.Close 'Close workbook - I intentionally leave the option for save\don't save\cancel
End Sub
