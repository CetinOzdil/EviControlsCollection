VERSION 5.00
Object = "{124D841B-26EA-4179-A897-6C656F2EB79B}#91.0#0"; "prjEviCollectionControl.ocx"
Begin VB.Form StylesDemo 
   Caption         =   "Styles Demo"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Different Colors"
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   6660
      Width           =   6540
      Begin prjEviCollectionControl.EviButton EviButton1 
         Height          =   720
         Index           =   1
         Left            =   300
         TabIndex        =   13
         Top             =   660
         Width           =   1000
         _ExtentX        =   1773
         _ExtentY        =   1270
         Caption         =   "OS && Red"
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
         ColorPressed    =   6776806
         ColorHover      =   9079551
         DefaultColors   =   0   'False
         BackColor       =   8421631
      End
      Begin prjEviCollectionControl.EviButton EviButton1 
         Height          =   720
         Index           =   2
         Left            =   3964
         TabIndex        =   14
         Top             =   660
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1270
         Caption         =   "XP && Orange"
         ButtonStyle     =   3
         ButtonStyleOS   =   0   'False
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
         ColorPressed    =   6793190
         ColorHover      =   9095935
         DefaultColors   =   0   'False
         BackColor       =   8438015
      End
      Begin prjEviCollectionControl.EviButton EviButton1 
         Height          =   720
         Index           =   4
         Left            =   1518
         TabIndex        =   15
         Top             =   660
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1270
         Caption         =   "W10 && Blue"
         ButtonStyle     =   5
         ButtonStyleOS   =   0   'False
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
         ColorPressed    =   15099751
         ColorHover      =   16747146
         DefaultColors   =   0   'False
         BackColor       =   16744576
      End
      Begin prjEviCollectionControl.EviButton EviButton1 
         Height          =   720
         Index           =   5
         Left            =   2741
         TabIndex        =   16
         Top             =   660
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1270
         Caption         =   "W7 && Green"
         ButtonStyle     =   6
         ButtonStyleOS   =   0   'False
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
         ColorPressed    =   6809191
         ColorHover      =   9109386
         DefaultColors   =   0   'False
         BackColor       =   8454016
      End
      Begin prjEviCollectionControl.EviButton EviButton1 
         Height          =   720
         Index           =   6
         Left            =   5190
         TabIndex        =   17
         Top             =   660
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1270
         Caption         =   "Std && Yellow"
         ButtonStyle     =   0
         ButtonStyleOS   =   0   'False
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
         ColorPressed    =   6809318
         ColorHover      =   9109503
         DefaultColors   =   0   'False
         BackColor       =   8454143
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Other Styles"
      Height          =   3465
      Index           =   1
      Left            =   3510
      TabIndex        =   8
      Top             =   2115
      Width           =   3150
      Begin prjEviCollectionControl.EviButton EviButton 
         Height          =   720
         Index           =   4
         Left            =   300
         TabIndex        =   9
         Top             =   495
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1270
         Caption         =   "Flat Button"
         ButtonStyle     =   1
         ButtonStyleOS   =   0   'False
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
      Begin prjEviCollectionControl.EviButton EviButton 
         Height          =   720
         Index           =   5
         Left            =   300
         TabIndex        =   10
         Top             =   1430
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1270
         Caption         =   "No Border Button"
         ButtonStyle     =   4
         ButtonStyleOS   =   0   'False
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
      Begin prjEviCollectionControl.EviButton EviButton 
         Height          =   720
         Index           =   6
         Left            =   300
         TabIndex        =   11
         Top             =   2355
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1270
         Caption         =   "Office 2010 Style"
         ButtonStyleOS   =   0   'False
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
   End
   Begin VB.Frame Frame 
      Caption         =   "All OS Styles"
      Height          =   4425
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2115
      Width           =   3150
      Begin prjEviCollectionControl.EviButton EviButton 
         Height          =   720
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   495
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1270
         Caption         =   "Standart (Win2K && Older)"
         ButtonStyle     =   0
         ButtonStyleOS   =   0   'False
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
      Begin prjEviCollectionControl.EviButton EviButton 
         Height          =   720
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   1430
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1270
         Caption         =   "Windows XP"
         ButtonStyle     =   3
         ButtonStyleOS   =   0   'False
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
      Begin prjEviCollectionControl.EviButton EviButton 
         Height          =   720
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   2365
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1270
         Caption         =   "Windows Vista / 7"
         ButtonStyle     =   6
         ButtonStyleOS   =   0   'False
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
      Begin prjEviCollectionControl.EviButton EviButton 
         Height          =   720
         Index           =   3
         Left            =   300
         TabIndex        =   7
         Top             =   3300
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1270
         Caption         =   "Windows 8 / 10"
         ButtonStyle     =   5
         ButtonStyleOS   =   0   'False
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Native Button && ButtonStyleOS"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   6540
      Begin prjEviCollectionControl.EviButton EviButton1 
         Height          =   720
         Index           =   0
         Left            =   3720
         TabIndex        =   2
         Top             =   660
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   1270
         Caption         =   "Evi Button"
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
      Begin VB.CommandButton Command 
         Caption         =   "Native Button"
         Height          =   720
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   660
         Width           =   2460
      End
   End
End
Attribute VB_Name = "StylesDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

