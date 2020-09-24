VERSION 5.00
Object = "{3BC4F112-323E-11D3-B079-BCE9DBE18C1B}#20.0#0"; "Counter.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Progress"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Interval        =   16
      Left            =   150
      Top             =   15
   End
   Begin TimingOCX.Speed Speed1 
      Height          =   765
      Left            =   555
      TabIndex        =   1
      Top             =   1170
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1349
      Interval        =   500
      MaxSpeed        =   500
      Caption         =   "Recs/Min"
   End
   Begin TimingOCX.Counter Counter1 
      Height          =   270
      Left            =   540
      TabIndex        =   0
      Top             =   420
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   476
      CharacterExtraY =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Digits          =   5
   End
   Begin VB.Label Label1 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   " Records processed "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1665
      TabIndex        =   2
      Top             =   420
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Speed1_QueryDistance(Mileage As Long)

    Mileage = Counter1 * 60

End Sub

Private Sub Timer1_Timer()

    Counter1 = Counter1 + 1 / 16

End Sub

':) Ulli's VB Code Formatter V2.11.3 (11.04.2002 14:31:33) 1 + 14 = 15 Lines
