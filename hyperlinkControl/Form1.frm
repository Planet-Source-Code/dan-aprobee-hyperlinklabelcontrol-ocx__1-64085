VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin projHyperlinkControl.Hyperlink Hyperlink4 
      Height          =   285
      Left            =   225
      TabIndex        =   3
      Top             =   1260
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   503
      Text            =   "Your Program Files"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      hyperlinkAddress=   "C:\Program Files"
   End
   Begin projHyperlinkControl.Hyperlink Hyperlink3 
      Height          =   285
      Left            =   225
      TabIndex        =   2
      Top             =   900
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   503
      Text            =   "PlanetSourceCode         (Java)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      hyperlinkAddress=   "http://www.planetsourcecode.com/vb/default.asp?lngWId=2"
      Image           =   "Form1.frx":0000
      ImageLeft       =   93
      ImageTop        =   0
      ImageWidth      =   20
      ImageHeight     =   20
   End
   Begin projHyperlinkControl.Hyperlink Hyperlink2 
      Height          =   285
      Left            =   225
      TabIndex        =   1
      Top             =   540
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   503
      Text            =   "PlanetSourceCode (C/C++)"
      HTextAlign      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      hyperlinkAddress=   "http://www.planetsourcecode.com/vb/default.asp?lngWId=3"
      Image           =   "Form1.frx":0512
      ImageLeft       =   133
      ImageTop        =   0
      ImageWidth      =   20
      ImageHeight     =   20
   End
   Begin projHyperlinkControl.Hyperlink Hyperlink1 
      Height          =   285
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   503
      Text            =   "PlanetSourceCode (VB)"
      HTextAlign      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Image           =   "Form1.frx":0A24
      ImageLeft       =   0
      ImageTop        =   0
      ImageWidth      =   20
      ImageHeight     =   20
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

