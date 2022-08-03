'------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Project: Spacious Storage Login & Permissions System
' Author: Adam Johnson, 2021
' Purpose: Add a layer of security being permissions, where only elevated permissions can view certain forms or buttons
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
Option Compare Database

Private Sub Form_Load()
    Me.tb_ID = Null
    Me.tb_pwd = Null
    Me.tb_ID.SetFocus
End Sub


Private Sub Login_BTN_Click()
    If IsNull(Me.tb_ID) Or IsNull(Me.tb_pwd) Then
      MsgBox "You must enter password and login ID.", vbOKOnly + vbInformation, "Required Data"
        Me.tb_ID.SetFocus
        Exit Sub
    End If


    If Me.tb_pwd.Value = DLookup("[Password]", "EmployeeDB", "[Username]='" & Me.tb_ID.Value & "'") Then
        If DLookup("[Position]", "EmployeeDB", "[Username]='" & Me.tb_ID.Value & "'") = "Developer" Then
        Dim FirstName As String
        FirstName = Me.tb_ID.Value
            DoCmd.OpenForm "Developer"
            DoCmd.Close acForm, "Login", acSaveYes
            Forms("Developer").LabelName.Caption = "Welcome, " & FirstName
            Forms("Developer").LabelLoc.Caption = "OrganizationName"
        Else
            DoCmd.OpenForm "user_menu"
            DoCmd.Close acForm, "Login", acSaveYes
        End If
    Else
      MsgBox "Password or login ID incorrect. Please Try Again", vbOKOnly + vbExclamation, "Invalid Entry!"
      Me.tb_pwd.SetFocus
    End If
End Sub
