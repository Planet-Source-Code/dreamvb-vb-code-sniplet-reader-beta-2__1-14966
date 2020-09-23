VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Project Options......"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Editor"
      Height          =   2460
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   6585
      Begin VB.PictureBox BCol 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4470
         ScaleHeight     =   225
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   1755
         Width           =   540
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   4425
         TabIndex        =   14
         Top             =   1710
         Width           =   1020
      End
      Begin VB.PictureBox FCOL 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   3315
         ScaleHeight     =   225
         ScaleWidth      =   480
         TabIndex        =   12
         Top             =   1755
         Width           =   540
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   3270
         TabIndex        =   11
         Top             =   1710
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   3330
         ScaleHeight     =   630
         ScaleWidth      =   2115
         TabIndex        =   8
         Top             =   660
         Width           =   2175
         Begin VB.Label lblStyle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ABCDEFabcdef"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Left            =   60
            TabIndex        =   9
            Top             =   195
            Width           =   1440
         End
      End
      Begin VB.ListBox lstSize 
         Height          =   1620
         Left            =   2325
         TabIndex        =   6
         Top             =   660
         Width           =   795
      End
      Begin VB.ListBox lstFont 
         Height          =   1620
         Left            =   105
         TabIndex        =   5
         Top             =   660
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Back-Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4395
         TabIndex        =   13
         Top             =   1410
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2460
         TabIndex        =   10
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sample"
         Height          =   195
         Left            =   4110
         TabIndex        =   7
         Top             =   390
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fore-Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3300
         TabIndex        =   4
         Top             =   1410
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   105
         TabIndex        =   3
         Top             =   390
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1275
      TabIndex        =   1
      Top             =   2685
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   2685
      Width           =   1155
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCan_Click()
    Unload frmOptions
    Form1.Show
    
End Sub

Private Sub Command1_Click()
    Config.Font_Name = lstFont.Text
    Config.Font_Size = lstSize.Text
    Config.Fore_Colour = FCOL.BackColor
    Config.Back_Colour = BCol.BackColor
    WriteINIChanges
    Form1.Show
    Unload frmOptions
    
End Sub

Private Sub Command2_Click()
On Error Resume Next
    FCOL.BackColor = ShowColor(hwnd)
    lblStyle.ForeColor = FCOL.BackColor
    Config.Fore_Colour = FCOL.BackColor
    If Err Then Err.Clear
End Sub

Private Sub Command3_Click()
On Error Resume Next
    BCol.BackColor = ShowColor(hwnd)
    Picture1.BackColor = BCol.BackColor
    Config.Back_Colour = BCol.BackColor
    If Err Then Err.Clear
    
End Sub

Private Sub Form_Load()
Dim IFont As Integer
    For IFont = 1 To Screen.FontCount - 1
        lstFont.AddItem Screen.Fonts(IFont)
    Next
    
    lstSize.AddItem "8"
    lstSize.AddItem "9"
    lstSize.AddItem "10"
    lstSize.AddItem "12"
    lstSize.AddItem "14"
    lstSize.AddItem "16"
    lstSize.AddItem "18"
    lstSize.AddItem "24"
    IFont = 0
    
End Sub

Private Sub lstFont_Click()
    lblStyle.Font = lstFont.Text
    Config.Font_Name = lstFont.Text
    
End Sub

Private Sub lstSize_Click()
    lblStyle.FontSize = Val(lstSize.Text)
    lblStyle.Top = Picture1.Height / 2 - 150
    Config.Font_Size = lstSize.Text
    
End Sub
