VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Code Sniplet Reader Beta 2"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   315
      Left            =   3765
      TabIndex        =   4
      Top             =   2025
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code Sniplet Reader Beta 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   180
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "About.frx":0000
         Top             =   135
         Width           =   480
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Visual Basic Soure Code Reader and Keeps all your source code in one place. were and when you most need it"
      Height          =   585
      Left            =   210
      TabIndex        =   3
      Top             =   810
      Width           =   4620
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Writen and Designed by Ben Jones"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1230
      TabIndex        =   2
      Top             =   1380
      Width           =   2520
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload about
    Form1.Show
    
End Sub
