VERSION 5.00
Object = "{124D841B-26EA-4179-A897-6C656F2EB79B}#91.0#0"; "prjEviCollectionControl.ocx"
Begin VB.Form Sample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sample"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
   Icon            =   "Sample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin prjEviCollectionControl.EviTextBox EviTextBox6 
      Height          =   2055
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3625
      Locked          =   -1  'True
      Scrollbar       =   2
      Focus           =   -1  'True
   End
   Begin prjEviCollectionControl.EviFrame EviFrame9 
      Height          =   1575
      Left            =   5880
      TabIndex        =   26
      Top             =   4320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2778
      GradientColor2  =   0
      BorderColor     =   16777215
      BackColor       =   16777215
      GradientPosition=   8
      Theme           =   4
      Style           =   3
      Caption         =   "MaskedBox"
      Appearance      =   1
      Icon            =   "Sample.frx":038A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowIcon        =   -1  'True
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Sample.frx":0724
         Left            =   720
         List            =   "Sample.frx":073A
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1080
         Width           =   1935
      End
      Begin prjEviCollectionControl.EviMaskedBox EviMaskedBox3 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mask"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   495
      End
   End
   Begin prjEviCollectionControl.EviFrame EviFrame8 
      Height          =   1575
      Left            =   3000
      TabIndex        =   22
      Top             =   4320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2778
      GradientColor2  =   14735318
      BorderColor     =   16777215
      BackColor       =   16777215
      GradientPosition=   1
      Theme           =   2
      Style           =   2
      Caption         =   "MaskedBox"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin prjEviCollectionControl.EviTextBox EviTextBox5 
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Locked          =   -1  'True
         Text            =   "dd-mmm-yy"
      End
      Begin prjEviCollectionControl.EviMaskedBox EviMaskedBox2 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Mask"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
   End
   Begin prjEviCollectionControl.EviFrame EviFrame7 
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2778
      GradientColor1  =   8421631
      GradientColor2  =   128
      BorderColor     =   128
      BackColor       =   12632319
      GradientPosition=   0
      Theme           =   5
      Style           =   1
      Caption         =   "MaskedBox"
      ForeColor       =   14737632
      ShadowColor     =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin prjEviCollectionControl.EviTextBox EviTextBox4 
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Locked          =   -1  'True
         Text            =   "hh:mm AM/PM"
      End
      Begin prjEviCollectionControl.EviMaskedBox EviMaskedBox1 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mask"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   495
      End
   End
   Begin prjEviCollectionControl.EviFrame EviFrame6 
      Height          =   975
      Left            =   5880
      TabIndex        =   16
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      GradientColor2  =   14735318
      BorderColor     =   16118256
      BackColor       =   16118256
      GradientPosition=   0
      Theme           =   2
      Style           =   1
      Caption         =   "PasswordChar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin prjEviCollectionControl.EviTextBox EviTextBox3 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         PasswordChar    =   "*"
         Text            =   "EviTextBox3"
      End
   End
   Begin prjEviCollectionControl.EviFrame EviFrame5 
      Height          =   975
      Left            =   3000
      TabIndex        =   14
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      GradientColor2  =   16711680
      BorderColor     =   16777215
      BackColor       =   16777215
      GradientPosition=   8
      Theme           =   3
      Style           =   3
      Caption         =   "Focus TextBox"
      Icon            =   "Sample.frx":077E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowIcon        =   -1  'True
      Begin prjEviCollectionControl.EviTextBox EviTextBox2 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Text            =   "EviTextBox2"
         Focus           =   -1  'True
      End
   End
   Begin prjEviCollectionControl.EviFrame EviFrame4 
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      GradientColor2  =   12117984
      BorderColor     =   15529718
      BackColor       =   15529718
      GradientPosition=   0
      Theme           =   1
      Style           =   1
      Caption         =   "TextBox"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin prjEviCollectionControl.EviTextBox EviTextBox1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         Text            =   "EviTextBox1"
      End
   End
   Begin prjEviCollectionControl.EviFrame EviFrame3 
      Height          =   1455
      Left            =   5880
      TabIndex        =   9
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2566
      GradientColor2  =   128
      BorderColor     =   16777215
      BackColor       =   16777215
      GradientPosition=   8
      Theme           =   5
      Style           =   3
      Caption         =   "Java Progressbar"
      Icon            =   "Sample.frx":0B18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowIcon        =   -1  'True
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   720
         Top             =   240
      End
      Begin prjEviCollectionControl.EviButton EviButton3 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "&Test"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14803425
         ShowBorder      =   -1  'True
         BEVEL           =   0
      End
      Begin prjEviCollectionControl.EviProgressBar EviProgressBar3 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         BrushStyle      =   0
         Color           =   12937777
         Style           =   5
         Color2          =   12937777
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   1800
   End
   Begin prjEviCollectionControl.EviFrame EviFrame2 
      Height          =   1455
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2566
      GradientColor1  =   12632256
      GradientColor2  =   0
      BorderColor     =   0
      BackColor       =   14737632
      GradientPosition=   0
      Theme           =   4
      Style           =   1
      Caption         =   "Search Progressbar"
      ForeColor       =   16777215
      Icon            =   "Sample.frx":0EB2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowIcon        =   -1  'True
      Begin prjEviCollectionControl.EviButton EviButton2 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "&Test"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjEviCollectionControl.EviProgressBar EviProgressBar2 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         BrushStyle      =   0
         Color           =   12937777
         Style           =   2
         Color2          =   12937777
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2280
      Top             =   1560
   End
   Begin prjEviCollectionControl.EviFrame EviFrame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2566
      Caption         =   "Standard Progressbar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin prjEviCollectionControl.EviButton EviButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "&Test"
         ButtonStyle     =   3
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   8454016
      End
      Begin prjEviCollectionControl.EviProgressBar EviProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         BrushStyle      =   0
         Color           =   12937777
         Color2          =   12937777
      End
   End
   Begin prjEviCollectionControl.Line Line1 
      Height          =   60
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   106
   End
   Begin prjEviCollectionControl.EviLabel EviLabel2 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      Caption         =   "UPDATE"
   End
   Begin prjEviCollectionControl.EviLabel EviLabel1 
      Height          =   540
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   953
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColor1    =   0
      ShadowColor2    =   12632256
      Appearance      =   1
      Caption         =   "Evi Collection Control XP"
   End
End
Attribute VB_Name = "Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Dim B As Integer
Dim c As Integer
Dim BB As Boolean

Private Sub Combo1_Click()
    EviMaskedBox3.Mask = Combo1.Text
    
    If Combo1.Text = "Standard" Then
        EviMaskedBox3.Text = 12345
    End If
    If Combo1.Text = "Currency" Then
        EviMaskedBox3.Text = 12345
    End If
    If Combo1.Text = "hh:mm AM/PM" Or Combo1.Text = "hh:mm" Then
        EviMaskedBox3.Text = Time
    End If
    If Combo1.Text = "dd-mmm-yy" Or Combo1.Text = "dd-mmm-yyyy" Then
        EviMaskedBox3.Text = Date
    End If
End Sub

Private Sub EviButton1_Click()
    Timer1.Enabled = True
End Sub

Private Sub EviButton2_Click()
    If EviButton2.Caption = "&Test" Then
        Timer2.Enabled = True
        EviButton2.Caption = "&Stop"
    Else
        Timer2.Enabled = False
        Me.EviProgressBar2.Value = 0
        EviButton2.Caption = "&Test"
        B = 0
    End If
End Sub

Private Sub EviButton3_Click()
    Timer3.Enabled = True
End Sub

Private Sub Form_Load()
    EviMaskedBox1.Mask = EviTextBox4.Text
    EviMaskedBox1.Text = Time

    EviMaskedBox2.Mask = EviTextBox5.Text
    EviMaskedBox2.Text = Date

    EviTextBox6.Text = "If you found bug or have request control, you can email me on cdozdil@gmail.com"
End Sub

Private Sub Timer1_Timer()
    a = a + 1
    
    If a > Me.EviProgressBar1.Max Then
        Me.EviProgressBar1.Value = 0
        Timer1.Enabled = False
        a = 0
        Exit Sub
    End If
    
    Me.EviProgressBar1.Value = a
End Sub

Private Sub Timer2_Timer()
    If BB = False Then
        B = B + 1
        If B > Me.EviProgressBar2.Max Then
            BB = True
        End If
    Else
        B = B - 1
        If B = 0 Then
            BB = False
        End If
    End If
    
    Me.EviProgressBar2.Value = B
End Sub

Private Sub Timer3_Timer()
    c = c + 1
    If c > Me.EviProgressBar3.Max Then
        Me.EviProgressBar3.Value = 0
        Timer3.Enabled = False
        c = 0
        Exit Sub
    End If
    
    Me.EviProgressBar3.Value = c
End Sub

Private Sub Timer4_Timer()
    EviButton3.BackColor = IIf(EviButton3.BackColor = vbBlue, vbButtonFace, vbBlue)
    EviButton2.BackColor = IIf(EviButton2.BackColor = vbGreen, vbButtonFace, vbGreen)
    EviButton1.BackColor = IIf(EviButton1.BackColor = vbRed, vbButtonFace, vbRed)
End Sub
