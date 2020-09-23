VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Sniplet Reader Beta 2"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "Reader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   5955
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6906
            Picture         =   "Reader.frx":063A
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   465
      Width           =   7860
      Begin VB.TextBox txtCodeSize 
         Height          =   315
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   570
         Width           =   1020
      End
      Begin VB.TextBox txtCodeVer 
         Height          =   315
         Left            =   6630
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   165
         Width           =   930
      End
      Begin VB.TextBox txtCodeDate 
         Height          =   315
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   165
         Width           =   1020
      End
      Begin VB.TextBox txtCodeBy 
         Height          =   315
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   525
         Width           =   3330
      End
      Begin VB.TextBox txtCodeName 
         Height          =   315
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   165
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Index           =   4
         Left            =   4560
         TabIndex        =   9
         Top             =   630
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "VB Ver"
         Height          =   195
         Index           =   3
         Left            =   6075
         TabIndex        =   7
         Top             =   195
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   5
         Top             =   195
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code By"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   570
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code Name"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   1
         Top             =   195
         Width           =   840
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7635
      Top             =   5970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   14
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Reader.frx":0B6C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4410
      Left            =   0
      TabIndex        =   12
      Top             =   1515
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   7779
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " Current Code Tips"
      TabPicture(0)   =   "Reader.frx":0E5E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ImageList2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Bevel1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   " View Source Code"
      TabPicture(1)   =   "Reader.frx":1390
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(1)=   "Bevel2"
      Tab(1).Control(2)=   "txtCodeMain"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   " Visual Basic Web Links"
      TabPicture(2)   =   "Reader.frx":18C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "Label3(1)"
      Tab(2).Control(2)=   "Image1(0)"
      Tab(2).Control(3)=   "Image1(1)"
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(5)=   "Image1(2)"
      Tab(2).Control(6)=   "Label6"
      Tab(2).Control(7)=   "Image1(3)"
      Tab(2).Control(8)=   "Label7"
      Tab(2).Control(9)=   "Image1(4)"
      Tab(2).Control(10)=   "Image1(5)"
      Tab(2).Control(11)=   "Label9"
      Tab(2).Control(12)=   "Image1(6)"
      Tab(2).Control(13)=   "Label10"
      Tab(2).Control(14)=   "Image1(7)"
      Tab(2).Control(15)=   "Label11"
      Tab(2).Control(16)=   "Label4"
      Tab(2).Control(17)=   "Label8"
      Tab(2).Control(18)=   "Label3(0)"
      Tab(2).Control(19)=   "Label3(2)"
      Tab(2).Control(20)=   "Label3(3)"
      Tab(2).Control(21)=   "Label3(4)"
      Tab(2).Control(22)=   "Label3(5)"
      Tab(2).Control(23)=   "Label3(6)"
      Tab(2).Control(24)=   "Label3(7)"
      Tab(2).Control(25)=   "Label3(8)"
      Tab(2).Control(26)=   "Label3(9)"
      Tab(2).Control(27)=   "Label3(10)"
      Tab(2).Control(28)=   "Label3(11)"
      Tab(2).Control(29)=   "Label3(12)"
      Tab(2).Control(30)=   "Label3(13)"
      Tab(2).Control(31)=   "Label3(14)"
      Tab(2).Control(32)=   "Label3(15)"
      Tab(2).Control(33)=   "Timer1"
      Tab(2).Control(34)=   "Picture1"
      Tab(2).Control(35)=   "Picture5"
      Tab(2).ControlCount=   36
      TabCaption(3)   =   " VB Code Serach"
      TabPicture(3)   =   "Reader.frx":1E14
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label14"
      Tab(3).Control(1)=   "Web"
      Tab(3).Control(2)=   "Picture6"
      Tab(3).ControlCount=   3
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   -74700
         Picture         =   "Reader.frx":2356
         ScaleHeight     =   900
         ScaleWidth      =   7020
         TabIndex        =   52
         Top             =   390
         Width           =   7020
         Begin VB.ComboBox cboSite 
            Height          =   315
            Left            =   4410
            TabIndex        =   54
            Top             =   105
            Width           =   1620
         End
         Begin VB.TextBox txtserach 
            Height          =   330
            Left            =   2295
            TabIndex        =   53
            Top             =   105
            Width           =   2070
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Devlopers  Source Code Web Serach"
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   105
            TabIndex        =   56
            Top             =   60
            Width           =   1635
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Find It"
            Height          =   390
            Left            =   1935
            TabIndex        =   55
            Top             =   90
            Width           =   345
         End
         Begin VB.Image Image2 
            Height          =   285
            Left            =   6120
            Picture         =   "Reader.frx":16CA8
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   30
         Left            =   -75000
         ScaleHeight     =   30
         ScaleWidth      =   7890
         TabIndex        =   51
         Top             =   4395
         Width           =   7890
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   75
         Left            =   -135
         ScaleHeight     =   75
         ScaleWidth      =   8160
         TabIndex        =   50
         Top             =   4395
         Width           =   8160
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   75
         Left            =   -75000
         ScaleHeight     =   75
         ScaleWidth      =   7860
         TabIndex        =   49
         Top             =   4350
         Width           =   7860
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   7860
            Y1              =   45
            Y2              =   45
         End
      End
      Begin Project1.Bevel Bevel2 
         Height          =   3930
         Left            =   -74925
         TabIndex        =   47
         Top             =   390
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   6932
      End
      Begin VB.TextBox txtCodeMain 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3930
         Left            =   -74925
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   48
         Top             =   390
         Width           =   7725
      End
      Begin Project1.Bevel Bevel1 
         Height          =   4005
         Left            =   30
         TabIndex        =   46
         Top             =   330
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   7064
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2490
         Left            =   -69810
         ScaleHeight     =   2490
         ScaleWidth      =   1785
         TabIndex        =   42
         Top             =   810
         Width           =   1785
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   0
            ScaleHeight     =   495
            ScaleWidth      =   1815
            TabIndex        =   44
            Top             =   -15
            Width           =   1815
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Latest Code Ticker By Planet Source-Code"
               Height          =   450
               Left            =   15
               TabIndex        =   45
               Top             =   75
               Width           =   1830
            End
         End
         Begin SHDocVwCtl.WebBrowser Web2 
            Height          =   4140
            Left            =   -195
            TabIndex        =   43
            Top             =   -1305
            Width           =   2310
            ExtentX         =   4075
            ExtentY         =   7302
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
         End
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   885
         Top             =   3660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   18
         ImageHeight     =   18
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":1715E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":175A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":179E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":17E24
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":18266
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":186A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":18AEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":18F2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":1936E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":197B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":19BF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":1A034
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":1A476
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":1A8B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":1ACFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":1B13C
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Reader.frx":1B24E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin SHDocVwCtl.WebBrowser Web 
         Height          =   2940
         Left            =   -74925
         TabIndex        =   40
         Top             =   1380
         Width           =   7635
         ExtentX         =   13467
         ExtentY         =   5186
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   -74670
         Top             =   2850
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4005
         Left            =   30
         TabIndex        =   13
         Top             =   330
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   7064
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   15
         Left            =   -70650
         TabIndex        =   39
         Top             =   840
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "k"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   14
         Left            =   -70905
         TabIndex        =   38
         Top             =   750
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "n"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   13
         Left            =   -71175
         TabIndex        =   37
         Top             =   690
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   12
         Left            =   -71385
         TabIndex        =   36
         Top             =   690
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   -71640
         TabIndex        =   35
         Top             =   630
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   -71970
         TabIndex        =   34
         Top             =   735
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   -72225
         TabIndex        =   33
         Top             =   675
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   -72495
         TabIndex        =   32
         Top             =   720
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   7
         Left            =   -72750
         TabIndex        =   31
         Top             =   675
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   -73005
         TabIndex        =   30
         Top             =   615
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "l"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   5
         Left            =   -73275
         TabIndex        =   29
         Top             =   630
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   -73545
         TabIndex        =   28
         Top             =   675
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   -73800
         TabIndex        =   27
         Top             =   675
         Width           =   165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   -74055
         TabIndex        =   26
         Top             =   615
         Width           =   150
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -74580
         TabIndex        =   25
         Top             =   585
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB Thunder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -71565
         MouseIcon       =   "Reader.frx":1B690
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1335
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Planet Source Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74415
         MouseIcon       =   "Reader.frx":1BF5A
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   1335
         Width           =   1740
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CodeArchive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -71565
         MouseIcon       =   "Reader.frx":1C824
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   2250
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   7
         Left            =   -71895
         Picture         =   "Reader.frx":1D0EE
         Top             =   2250
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB Only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -71565
         MouseIcon       =   "Reader.frx":1D42A
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1950
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   6
         Left            =   -71895
         Picture         =   "Reader.frx":1DCF4
         Top             =   1950
         Width           =   255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -71565
         MouseIcon       =   "Reader.frx":1E030
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1665
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   5
         Left            =   -71895
         Picture         =   "Reader.frx":1E8FA
         Top             =   1665
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   4
         Left            =   -71895
         Picture         =   "Reader.frx":1EC36
         Top             =   1335
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DreamVb Home Page"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74415
         MouseIcon       =   "Reader.frx":1EF72
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   2250
         Width           =   1890
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   -74745
         Picture         =   "Reader.frx":1F83C
         Top             =   2250
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Microsoft Site"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74415
         MouseIcon       =   "Reader.frx":1FB78
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1950
         Width           =   1200
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   -74745
         Picture         =   "Reader.frx":20442
         Top             =   1950
         Width           =   255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB World"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   -74415
         MouseIcon       =   "Reader.frx":2077E
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1665
         Width           =   840
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   -74745
         Picture         =   "Reader.frx":21048
         Top             =   1665
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   -74745
         Picture         =   "Reader.frx":21384
         Top             =   1335
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   -74265
         TabIndex        =   16
         Top             =   630
         Width           =   75
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFEFE7&
         Height          =   4020
         Left            =   -74970
         TabIndex        =   15
         Top             =   345
         Width           =   7800
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808080&
         Height          =   990
         Left            =   -74925
         TabIndex        =   41
         Top             =   345
         Width           =   7770
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   688
      ButtonWidth     =   661
      ButtonHeight    =   635
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add New Tips"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Extract Tip"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy Tip"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Config Options"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Project Builder"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Project Extracter"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Run Tip"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   -135
      X2              =   1590
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   -135
      X2              =   1590
      Y1              =   435
      Y2              =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Db As Database
'Dim RecSet As Recordset

Dim Code_Name As Collection
Dim Code_Date As Collection
Dim Code_Size As Collection
Dim Code_Ver As Collection
Dim Code_By As Collection
Dim Vb_Code As Collection

Private Sub Form_Load()
Dim EstTime As Single
Dim Icount As Long
Dim A1, A2, A3, A4, A5, A6 As String

    TimeOpen = Timer
    CenterForm Form1
    ReadWriteINI
    CodeDBPath = AddBackSlash(App.Path) & "VbCode.mdb"
    If FileExists(CodeDBPath) = 0 Then
        MsgBox "Can't Find Database", vbCritical
        Unload Form1: End
        Exit Sub
    Else
        Set Code_Name = New Collection
        Set Code_Date = New Collection
        Set Code_Size = New Collection
        Set Code_Ver = New Collection
        Set Code_By = New Collection
        Set Vb_Code = New Collection
        Set Db = OpenDatabase(CodeDBPath)
        
        With Db
            Set RecSet = .OpenRecordset("Code")
            If RecSet.RecordCount = 0 Then
                Exit Sub
            Else
                With RecSet
                .Edit
                    For Icount = 1 To RecSet.RecordCount
                    On Error Resume Next
                        ListView1.ListItems.Add Icount, "", !CodeName, 0, 1
                        A1 = !CodeName
                        A2 = !CodeDate
                        A3 = !CodeSize
                        A4 = !CodeVer
                        A5 = !CodeBy
                        A6 = !MainCode
                        Code_Name.Add A1
                        Code_Date.Add A2
                        Code_Size.Add A3
                        Code_Ver.Add A4
                        Code_By.Add A5
                        Vb_Code.Add A6
                    .MoveNext
                    Next
                End With
            End If
        End With
    End If
    
    txtCodeMain.BackColor = Config.Back_Colour
    txtCodeMain.ForeColor = Config.Fore_Colour
    txtCodeMain.FontName = Config.Font_Name
    txtCodeMain.FontSize = Config.Font_Size
    
    Web.Navigate "about:"
    Web2.Navigate "http://www.planet-source-code.com/vb/LinkToUs/ScrollingCode.asp?lngWId=1&"
    
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    
    FlatBorder txtCodeMain.hWnd
    
    StatusBar1.Panels(1).Text = "Total tips " & RecSet.RecordCount
    cboSite.AddItem "Please Select One"
    cboSite.AddItem "CodeAchive"
    cboSite.AddItem "Visual Basic Explorer"
    cboSite.AddItem "Visual Basic Code Room"
    cboSite.AddItem "VB World"
    cboSite.ListIndex = 0
    Icount = 0
    A1 = "": A2 = "": A3 = "": A4 = "": A5 = "": A6 = ""
    
    If Module1.FolderExists(AddBackSlash(App.Path) & "Run") = 0 Then
        MkDir AddBackSlash(App.Path) & "Run"
    End If
        
    StatusBar1.Panels(2).Text = "DB Open: " & Format(Timer - TimeOpen, "0.00") & " Seconds"
    
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = Form1.ScaleWidth - 1
    Line1(1).X2 = Form1.ScaleWidth - 1
    
End Sub

Private Sub Image2_Click()
    If Len(txtserach.Text) = 0 Then
        MsgBox "You must enter a serach string eg Bitmap,Games,Ocx etc", vbInformation
    ElseIf cboSite.ListIndex = 0 Then
        MsgBox "You must select a Item from the list", vbInformation
    Else
        SerachSites cboSite.ListIndex, Trim(txtserach.Text), Web
End If

End Sub

Private Sub Label10_Click()
    RunProgram hWnd, "http://www.vbonly.com", vsNormal

End Sub

Private Sub Label11_Click()
    RunProgram hWnd, "http://www.codearchive.com", vsNormal

End Sub

Private Sub Label4_Click()
    RunProgram hWnd, "http://www.planet-source-code.com/vb", vsNormal
    
End Sub

Private Sub Label5_Click()
        RunProgram hWnd, "http://www.vbworld.com", vsNormal

End Sub

Private Sub Label6_Click()
        RunProgram hWnd, "http://www.microsoft.com/vbasic", vsNormal

End Sub

Private Sub Label7_Click()
        RunProgram hWnd, "http://www.codearchive.com/~dreamvb", vsNormal

End Sub

Private Sub Label8_Click()
            RunProgram hWnd, "http://www.vbthunder.com", vsNormal

End Sub

Private Sub Label9_Click()
            RunProgram hWnd, "http://www.maxcode.com", vsNormal
    
End Sub

Private Sub ListView1_Click()
    txtCodeName.Text = Code_Name(ListView1.SelectedItem.Index)
    txtCodeDate.Text = Code_Date(ListView1.SelectedItem.Index)
    txtCodeSize.Text = Format(Val(Trim(Code_Size(ListView1.SelectedItem.Index))), "#,#") & " Bytes"
    txtCodeVer.Text = Code_Ver(ListView1.SelectedItem.Index)
    txtCodeBy.Text = Code_By(ListView1.SelectedItem.Index)
    txtCodeMain.Text = Vb_Code(ListView1.SelectedItem.Index)
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(8).Enabled = True
    
    Extra = txtCodeName.Text
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Caption = " Visual Basic Web Links" Then
        Timer1.Enabled = True
        Timer1.Interval = 100
    Else
        Timer1.Interval = 0
        Timer1.Enabled = False
    End If
    
End Sub

Private Sub Timer1_Timer()
Static I As Integer
    On Error Resume Next
    I = I + 1
    If I = Label3.Count Then I = 0: Label3(15).FontBold = False: Label3(15).ForeColor = vbBlack
        Label3(I).FontBold = True
        Label3(I).ForeColor = vbRed
        Label3(I - 1).ForeColor = vbBlack
        Label3(I - 1).FontBold = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            Form1.Hide
            Form2.Show
        Case 2
            Form1.Hide
            Form3.Show
            
        Case 4
            Clipboard.Clear
            Clipboard.SetText txtCodeMain.Text
            MsgBox "Text has been copyed to the clipbaord", vbInformation
        Case 5
            Form1.Hide
            frmOptions.Show
        Case 8
            Open AddBackSlash(App.Path) & "Run\" & Extra & ".bas" For Output As #1
                Print #1, , "Attribute VB_Name =" & Chr(34) & "VBCode32" & Chr(34)
                Print #1, , txtCodeMain.Text
            Close #1
            RunProgram hWnd, AddBackSlash(App.Path) & "Run\" & Extra & ".bas", vsNormal
        Case 10
            Form1.Hide
            about.Show
        Case 11
            ans = _
            MsgBox("Do you want to quit now", _
            vbYesNo, "Quit Program")
            If ans = vbNo Then
                Exit Sub
            Else
                Unload Form1: End
            End If
                
        End Select
        
End Sub



