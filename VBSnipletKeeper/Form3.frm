VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extract Tip"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   1200
      TabIndex        =   2
      Top             =   1530
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Extarct"
      Height          =   350
      Left            =   30
      TabIndex        =   1
      Top             =   1530
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   1380
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   4425
      Begin VB.OptionButton OptBas 
         Caption         =   "Extract as Visual Basic Module"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   570
         Width           =   4005
      End
      Begin VB.OptionButton OptText 
         Caption         =   "Extract as Text Document"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   4005
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mPath As String

Private Sub Command1_Click()
Dim File1, File2 As Long
    File1 = FreeFile
    File2 = FreeFile
    mPath = AddBackSlash(mPath)
    If OptText Then
        Open mPath & Extra & ".txt" For Output As #File1
            Print #File1, , Form1.txtCodeMain.Text
        Close #File1
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        MsgBox "Your tip has been saved to " & mPath & Extra
        Unload Form3: Form1.Show
        ElseIf OptBas Then
            Open mPath & Extra & ".bas" For Output As #File2
                Print #File2, , "Attribute VB_Name = " & Chr(34) & "Module1" & Chr(34)
                Print #File2, , Form1.txtCodeMain.Text
            Close #hfile
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        MsgBox "Your tip has been saved to " & mPath & Extra
        Unload Form3: Form1.Show
        Exit Sub
    Else
        MsgBox "You must select a extract option", vbInformation
    End If
    
End Sub

Private Sub Command2_Click()
    Unload Form3
    Form1.Show
    
End Sub

Private Sub Form_Load()
    If FolderExists(AddBackSlash(App.Path) & "Tips") = 0 Then
        MkDir AddBackSlash(App.Path) & "Tips"
    End If
    mPath = AddBackSlash(App.Path) & "Tips"
    
End Sub
