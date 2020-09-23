VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Code Tip"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Finished"
      Height          =   350
      Left            =   6285
      TabIndex        =   17
      Top             =   5790
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Save Tip"
      Height          =   350
      Left            =   5295
      TabIndex        =   16
      Top             =   5790
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   4305
      TabIndex        =   15
      Top             =   5790
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   300
      Left            =   4695
      TabIndex        =   14
      Top             =   2085
      Width           =   390
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1605
      TabIndex        =   13
      Top             =   2085
      Width           =   3045
   End
   Begin VB.ComboBox cboVer 
      Height          =   315
      Left            =   1155
      TabIndex        =   11
      Top             =   1410
      Width           =   1755
   End
   Begin VB.TextBox txtVBCode 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   2685
      Width           =   7260
   End
   Begin VB.TextBox txt4 
      Height          =   285
      Left            =   1155
      TabIndex        =   9
      Top             =   1125
      Width           =   1125
   End
   Begin VB.TextBox txt3 
      Height          =   285
      Left            =   1155
      TabIndex        =   8
      Top             =   810
      Width           =   1110
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   1155
      TabIndex        =   7
      Top             =   510
      Width           =   2955
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   1155
      TabIndex        =   6
      Top             =   210
      Width           =   2955
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Load Tips form File"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type in your code below or use the Option below to load in text files"
      Height          =   195
      Index           =   5
      Left            =   75
      TabIndex        =   5
      Top             =   1725
      Width           =   4770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code Verision"
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   4
      Top             =   1395
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code Size"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   1095
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code Date"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code By"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   510
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code Name"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   210
      Width           =   840
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Filenum As Long
Dim lzFilename, FileExt As String
    Filenum = FreeFile
    lzFilename = OpenFile("All Files(*.txt Text Files)" + Chr$(0) + "*.txt")
    FileExt = UCase(Right(lzFilename, 3))
    If Len(FileExt) = 0 Then
        Exit Sub
        ElseIf FileExt <> "TXT" Then
            MsgBox "This is not a viald Text Doucement Filename", vbExclamation
            Exit Sub
        Else
            txtFile.Text = lzFilename
            Open lzFilename For Input As #Filenum
                txtVBCode.Text = Input(LOF(Filenum), Filenum)
            Close #Filenum
    End If
    lzFilename = "": FileExt = ""
    
End Sub

Private Sub Command3_Click()
    If txt1 = "" Then
        MsgBox "You must enter the name of the code", vbInformation
        ElseIf txt2 = "" Then
            MsgBox "You must enter who the code is by", vbInformation
        ElseIf txt3 = "" Then
            MsgBox "You must enter the date the code was created eg (DD-MM-YY)", vbInformation
        ElseIf txt4 = "" Then
            MsgBox "You have not enteded the size of the code", vbInformation
        ElseIf cboVer.ListIndex = 0 Then
            MsgBox "You must selected the codes verision", vbInformation
        ElseIf txtVBCode = "" Then
            MsgBox "You must enter some code", vbInformation
        Else
            Set Db = OpenDatabase(CodeDBPath)
            With Db
                Set RecSet = .OpenRecordset("Code")
                If RecSet.RecordCount = 0 Then
                    Exit Sub
                Else
                    With RecSet
                    .AddNew
                        !CodeName = Trim(txt1)
                        !CodeDate = Trim(txt3)
                        !CodeSize = Trim(txt4)
                        !CodeVer = cboVer.Text
                        !CodeBy = Trim(txt2)
                        !MainCode = txtVBCode
                    .Update
                    End With
                End If
            End With
    End If
    
    MsgBox "You new Tip have been added to the database", vbInformation
    txt1 = ""
    txt2 = ""
    txt3 = ""
    txt4 = ""
    cboVer.ListIndex = 0
    txtVBCode = ""
    
End Sub

Private Sub Command4_Click()
    Unload Form2
    Form1.Show
    
End Sub

Private Sub Form_Load()

    If FileExists(CodeDBPath) = 0 Then
        MsgBox "can t find " & CodeDBPath
        Unload Form1: End
    End If
    
    CenterForm Form2
    cboVer.AddItem "Please Select One"
    cboVer.AddItem "VB 2.0"
    cboVer.AddItem "VB 3.0"
    cboVer.AddItem "VB 4/16"
    cboVer.AddItem "VB 4.0"
    cboVer.AddItem "VB 5.0"
    cboVer.AddItem "V6.0"
    cboVer.AddItem "VB7.NET"
    cboVer.ListIndex = 0
    
End Sub

Private Sub Image1_Click()

End Sub
