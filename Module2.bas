Attribute VB_Name = "Module2"

Sub CreateAddInMenu()
    On Error Resume Next
    Dim customMenu As CommandBarControl
    Dim checkButton As CommandBarControl

    ' Remove existing menu if it already exists
    Application.CommandBars("Menu Bar").Controls("������ ������").Delete

    ' Add a new menu to the Menu Bar
    Set customMenu = Application.CommandBars("Menu Bar").Controls.add(Type:=msoControlPopup, Temporary:=True)
    customMenu.Caption = "������ ������"

    ' Add a Check button under this new menu
    Set checkButton = customMenu.Controls.add(Type:=msoControlButton)
    checkButton.Caption = "�������"
    checkButton.OnAction = "ShowUserForm"  ' Calls ShowUserForm macro
End Sub

Sub AutoExec()
    ' Automatically create menu on startup
    CreateAddInMenu
End Sub



Sub ShowUserForm()
    ' Show the UserForm for input
    UserFormCheck.Show vbModeless
End Sub

