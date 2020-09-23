Attribute VB_Name = "Module1"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200

Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Private Const HWND_NOTOPMOST = -2




Public TimeOpen As Single
Public CodeDBPath As String
Public Db As Database
Public RecSet As Recordset
Public Extra As String
Public Config As Conf


Enum WinShow
    vsHide = 0
    vsNormal = 1
    vsMinSized = 2
    vsMaxSized = 3
End Enum

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type Conf
    Font_Name As String
    Font_Size As String
    Fore_Colour As String
    Back_Colour As String
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Sub FlatBorder(ByVal hWnd As Long)
Dim TFlat As Long
    TFlat = GetWindowLong(hWnd, GWL_EXSTYLE)
    TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    SetWindowLong hWnd, GWL_EXSTYLE, TFlat
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Sub
Function WriteINIChanges()
    SaveSetting "VbCode32", "Config", "FontName", Config.Font_Name
    SaveSetting "VbCode32", "Config", "FontSize", Config.Font_Size
    SaveSetting "VbCode32", "Config", "ForeColour", Config.Fore_Colour
    SaveSetting "VbCode32", "Config", "BackColour", Config.Back_Colour
    
End Function
Function ReadWriteINI()
    Config.Font_Name = GetSetting("VbCode32", "Config", "FontName")
    Config.Font_Size = GetSetting("VbCode32", "Config", "FontSize")
    Config.Fore_Colour = GetSetting("VbCode32", "Config", "ForeColour")
    Config.Back_Colour = GetSetting("VbCode32", "Config", "BackColour")
    
    If Config.Font_Name = "" Then
        SaveSetting "VbCode32", "Config", "FontName", "Courier New"
    End If

    If Config.Font_Size = "" Then
        SaveSetting "VbCode32", "Config", "FontSize", "10"
    End If

    If Config.Fore_Colour = "" Then
        SaveSetting "VbCode32", "Config", "ForeColour", "0"
    End If

    If Config.Back_Colour = "" Then
        SaveSetting "VbCode32", "Config", "BackColour", "16777215"
    End If
    
End Function
Public Function ShowColor(Handle As Long) As Long
Dim TCol As CHOOSECOLOR
Dim Custcolor(41) As Long
Dim lReturn As Long
    
    TCol.lStructSize = Len(TCol)
    TCol.hwndOwner = Handle
    TCol.hInstance = App.hInstance
    TCol.lpCustColors = StrConv(CustomColors, vbUnicode)
    TCol.flags = 0
    
    If CHOOSECOLOR(TCol) <> 0 Then
        ShowColor = TCol.rgbResult
        CustomColors = StrConv(TCol.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If

End Function
Private Function RemoveNulls(lzString As String) As String
Dim Xpos As Integer
    Xpos = InStr(lzString, vbNullChar)
    If Xpos > 0 Then
        lzString = Left(lzString, Len(lzString) - 1)
        RemoveNulls = lzString
    End If
    
End Function
Public Function OpenFile(Pattern As String) As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = Pattern
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Open Text File"
        ofn.flags = 0
        
        a = GetOpenFileName(ofn)
        If (a) Then
                OpenFile = RemoveNulls(Trim(ofn.lpstrFile))
        End If
        
 End Function
Public Function OpenMid() As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files(*.mid Multmedia Files)" + Chr$(0) + "*.mid"
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Mid Files"
        ofn.flags = 0
        
        a = GetOpenFileName(ofn)
        If (a) Then
                OpenMid = RemoveNulls(Trim(ofn.lpstrFile))
        End If
        
 End Function
Public Function RunProgram(mHwnd As Long, ProgramNamePath As String, ShowWindow As WinShow)
    ShellExecute mHwnd, vbNullString, ProgramNamePath, vbNullString, vbNullString, ShowWindow
    
End Function
Function CenterForm(Frm As Form)
    With Frm
        .Top = (Screen.Height - Frm.Height) / 2
        .Left = (Screen.Width - Frm.Width) / 2
    End With
    
End Function
Public Function FolderExists(ByVal Foldername As String) As Integer
    If Dir(Foldername, vbDirectory) = "" Then FolderExists = 0 Else FolderExists = 1
    
End Function
Public Function FileExists(ByVal Filename As String) As Integer
    If Dir(Filename) = "" Then FileExists = 0 Else FileExists = 1
    
End Function
Function SerachSites(SiteOp As Integer, SerachFor As String, WebWin As WebBrowser)
Dim Site1, Site2, Site3, Site4 As String
    
    Site1 = "http://www.codearchive.com/search/search.cgi?startat=0&search=" & SerachFor & "&section=VB"
    Site2 = "http://search.atomz.com/search/?sp-a=000700ff-sp00000000&sp-q=" & SerachFor
    Site3 = "http://www.vbcode.com/asp/code.asp?SortBy=&lstCategory=&KeywordSearch=" & SerachFor & "&SearchType=ExactPhrase&intpage=1"
    Site4 = "http://www.vb-world.net/cgi-bin/searchredir.cgi?search=" & SerachFor & "&whereto=VBWORLD"
    
    Select Case SiteOp
        Case 1
            WebWin.Navigate Site1
        Case 2
            WebWin.Navigate Site2
        Case 3
            WebWin.Navigate Site3
        Case 4
            WebWin.Navigate Site4
        End Select
        
End Function
Function AddBackSlash(Pathname As String) As String
Dim TBackSlash As String

    If Not Right(Pathname, 1) = "\" Then
        TBackSlash = Pathname & "\"
    Else
        TBackSlash = Pathname
    End If
    AddBackSlash = TBackSlash
    
End Function
