              Attribute VB_Name = "Module1"
              ' How to Disable the X Button on a form
' Start a new project and add a new module Project >> Add module click open
' Place the code below into the new module


Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Function RemoveXFormMenu(mHwnd As Long)
Dim H_Menu As Long
    H_Menu = GetSystemMenu(mHwnd, 0)
    
     If H_Menu < 0 Then
        MsgBox "Can't find menu", vbCritical
        Exit Function
      Else
        RemoveMenu H_Menu, 6, MF_BYPOSITION
        RemoveMenu H_Menu, 5, MF_BYPOSITION
    End If
    
End Function

' Now Place this code into the form_Load Event and press 5F
' You will see that the X button on your form is now disabled

 Module1.RemoveXFormMenu Form1.hwnd
