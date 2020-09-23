VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matrix Settings"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Colour Options"
      Height          =   1335
      Left            =   3000
      TabIndex        =   16
      Top             =   840
      Width           =   2175
      Begin VB.CheckBox Check2 
         Caption         =   "Fading(Slower)"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox Check4 
         Caption         =   "HighLights(Slower)"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Value           =   1  'Checked
         Width           =   1740
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Multiple Colours(Slower)"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Value           =   1  'Checked
         Width           =   1980
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   615
      Left            =   3000
      TabIndex        =   14
      Top             =   120
      Width           =   2175
      Begin VB.CheckBox Check1 
         Caption         =   "From Top"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resolution"
      Height          =   855
      Left            =   3000
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   "1024"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drop Options"
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
      Begin MSComctlLib.Slider Slider1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   510
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Min             =   10
         Max             =   100
         SelStart        =   100
         TickStyle       =   3
         Value           =   100
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1215
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   150
         SelStart        =   30
         TickStyle       =   3
         Value           =   30
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   1000
         SelStart        =   20
         TickStyle       =   3
         Value           =   20
      End
      Begin MSComctlLib.Slider Slider4 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   2580
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Min             =   1
         SelStart        =   4
         TickStyle       =   3
         Value           =   4
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Maximum Drop Length"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Wait Before Clearing"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Number of Dropping Columns"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   2070
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fading Speed"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   990
      End
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Slider4.Enabled = True
    Else
        Slider4.Enabled = False
    End If
End Sub


Sub SaveSets()
    Call SaveSetting("Kevin Pfister's Matrix", "Drops", "MaxDrop", Slider1.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Drops", "BeforeClean", Slider2.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Drops", "DropsRunning", Slider3.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Drops", "FadeSpeed", Slider4.Value)
    
    Call SaveSetting("Kevin Pfister's Matrix", "Options", "FromTop", Check1.Value)
    
    Call SaveSetting("Kevin Pfister's Matrix", "Colour", "Fade", Check2.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Colour", "MColours", Check3.Value)
    Call SaveSetting("Kevin Pfister's Matrix", "Colour", "HighLights", Check4.Value)
    
    Call SaveSetting("Kevin Pfister's Matrix", "Res", "XY", Combo1)
    
End Sub

Private Sub CmdAbout_Click()
    MsgBox ("Matrix ScreenSaver, Created By Kevin Pfister")
End Sub

Private Sub CmdCancel_Click()
    Call SaveSets
    End
End Sub

Private Sub CmdOk_Click()
    Call SaveSets
    End
End Sub

Private Sub Form_Load()
    Slider1.Value = GetSetting("Kevin Pfister's Matrix", "Drops", "MaxDrop", 100)
    Slider2.Value = GetSetting("Kevin Pfister's Matrix", "Drops", "BeforeClean", 30)
    Slider3.Value = GetSetting("Kevin Pfister's Matrix", "Drops", "DropsRunning", 20)
    Slider4.Value = GetSetting("Kevin Pfister's Matrix", "Drops", "FadeSpeed", 4)
    
    Check1.Value = GetSetting("Kevin Pfister's Matrix", "Options", "FromTop", 1)
    
    Check2.Value = GetSetting("Kevin Pfister's Matrix", "Colour", "Fade", 1)
    Check3.Value = GetSetting("Kevin Pfister's Matrix", "Colour", "MColours", 1)
    Check4.Value = GetSetting("Kevin Pfister's Matrix", "Colour", "HighLights", 1)
    
    Combo1 = GetSetting("Kevin Pfister's Matrix", "Res", "XY", "1024 by 768 pixels")
    
    Combo1.AddItem ("640 by 480 pixels")
    Combo1.AddItem ("720 by 480 pixels")
    Combo1.AddItem ("720 by 576 pixels")
    Combo1.AddItem ("800 by 600 pixels")
    Combo1.AddItem ("1024 by 768 pixels")
    Combo1.AddItem ("1152 by 864 pixels")
    Combo1.AddItem ("1280 by 960 pixels")
    Combo1.AddItem ("1280 by 1024 pixels")
    Combo1.AddItem ("1600 by 1200 pixels")
End Sub

