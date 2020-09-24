VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl Speed 
   CanGetFocus     =   0   'False
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ForwardFocus    =   -1  'True
   PropertyPages   =   "Speed.ctx":0000
   ScaleHeight     =   735
   ScaleWidth      =   2175
   ToolboxBitmap   =   "Speed.ctx":001B
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   210
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Speed Display"
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2370
      Top             =   30
   End
   Begin VB.Label lbCapt 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1035
      TabIndex        =   9
      Top             =   0
      Width           =   105
   End
   Begin VB.Label lbTick 
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      Height          =   195
      Index           =   4
      Left            =   1575
      TabIndex        =   8
      Top             =   390
      Width           =   30
   End
   Begin VB.Label lbTick 
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      Height          =   195
      Index           =   3
      Left            =   540
      TabIndex        =   7
      Top             =   390
      Width           =   30
   End
   Begin VB.Label lbNum 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   2
      Left            =   2130
      TabIndex        =   6
      Top             =   570
      Width           =   45
   End
   Begin VB.Label lbNum 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   1
      Left            =   1035
      TabIndex        =   5
      Top             =   570
      Width           =   60
   End
   Begin VB.Label lbNum 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   570
      Width           =   90
   End
   Begin VB.Label lbTick 
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      Height          =   195
      Index           =   2
      Left            =   2115
      TabIndex        =   3
      Top             =   390
      Width           =   30
   End
   Begin VB.Label lbTick 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      Height          =   195
      Index           =   1
      Left            =   1050
      TabIndex        =   2
      Top             =   390
      Width           =   30
   End
   Begin VB.Label lbTick 
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      Height          =   195
      Index           =   0
      Left            =   15
      TabIndex        =   1
      Top             =   390
      Width           =   30
   End
End
Attribute VB_Name = "Speed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefLng A-Z
Private myEnabled    As Boolean
Private Virgin       As Boolean
Private Diff(0 To 5) As Long
Private ubDiff       As Long
Private PrevValue    As Long
Public Event QueryDistance(Mileage As Long)

Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Sets / Returns the interval in milliseconds between QueryDistance events."
Attribute Interval.VB_HelpID = 20004
Attribute Interval.VB_ProcData.VB_Invoke_Property = ";Verhalten"

    Interval = tmr.Interval

End Property

Public Property Let Interval(ByVal nwInterval As Long)

    If nwInterval >= 10 Then
        tmr.Interval() = NextValid(nwInterval)
        PropertyChanged "Interval"
      Else 'NOT NWINTERVAL...
        Err.Raise 380
    End If

End Property

Public Property Get MaxSpeed() As Single
Attribute MaxSpeed.VB_Description = "Ses / Returns the Maximum Scale Value. Rounded to the next 1 - 2 - 5 value."
Attribute MaxSpeed.VB_HelpID = 20005
Attribute MaxSpeed.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute MaxSpeed.VB_UserMemId = 0

    MaxSpeed = pgb.Max

End Property

Public Property Let MaxSpeed(ByVal nwMax As Single)

    pgb.Max() = NextValid(nwMax)
    lbNum(2) = pgb.Max
    lbNum(1) = pgb.Max / 2
    PropertyChanged "MaxSpeed"

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets / Returns the caption for the Control."
Attribute Caption.VB_HelpID = 20000
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"

    Caption = lbCapt

End Property

Public Property Let Caption(ByVal nwCaption As String)

    lbCapt = nwCaption

End Property

Public Property Get CurrentSpeed() As Long
Attribute CurrentSpeed.VB_Description = "Sets / Returs the current speed display value."
Attribute CurrentSpeed.VB_HelpID = 20001
Attribute CurrentSpeed.VB_ProcData.VB_Invoke_Property = ";Daten"

    CurrentSpeed = pgb.Value

End Property

Public Property Let CurrentSpeed(ByVal nwValue As Long)

    If nwValue >= 0 And nwValue <= pgb.Max Then
        pgb.Value() = nwValue
        PropertyChanged "CurrentSpeed"
      Else 'NOT NWVALUE...
        Err.Raise 380
    End If

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets / Returns wether the Control will fire QueryDistance events."
Attribute Enabled.VB_HelpID = 20002
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Verhalten"

    Enabled = myEnabled

End Property

Public Property Let Enabled(ByVal nwEnabled As Boolean)

    myEnabled = nwEnabled
    PropertyChanged "Enabled"
    tmr.Enabled = myEnabled And Ambient.UserMode

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets / Returns the forecolor for the Control."
Attribute ForeColor.VB_HelpID = 20003
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"

    ForeColor = lbCapt.ForeColor

End Property

Public Property Let ForeColor(ByVal nwColor As OLE_COLOR)

  Dim i As Long

    For i = 0 To 2
        lbTick(i).ForeColor = nwColor
        lbNum(i).ForeColor = nwColor
    Next i
    lbTick(i).ForeColor = nwColor
    lbTick(i + 1).ForeColor = nwColor
    lbCapt.ForeColor = nwColor

End Property

Public Property Get Font() As Font

    Set Font = lbCapt.Font

End Property

Public Property Set Font(ByVal nwFont As Font)

    Set lbCapt.Font = nwFont
    lbNum(0).Font.Name = nwFont.Name
    lbNum(1).Font.Name = nwFont.Name
    lbNum(2).Font.Name = nwFont.Name

End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)

    UserControl.BackColor = Ambient.BackColor

End Sub

Private Sub UserControl_Initialize()

  Dim i As Long

    ubDiff = UBound(Diff)
    For i = 1 To ubDiff
        Diff(i) = 0
    Next i
    PrevValue = 0
    pgb.Min = 0
    MaxSpeed = 1000

End Sub

Private Sub UserControl_InitProperties()

    Virgin = True
    Enabled = True

End Sub

Private Sub UserControl_Resize()

  Dim i As Long

    pgb.Width = UserControl.Width
    lbCapt.Left = (pgb.Width - lbCapt.Width) / 2
    lbTick(2).Left = pgb.Width - 60
    lbTick(1).Left = (lbTick(0).Left + lbTick(2).Left) / 2
    lbNum(1).Left = lbTick(1).Left + (lbTick(1).Width - lbNum(1).Width) / 2
    lbNum(2).Left = pgb.Width - lbNum(2).Width
    On Error Resume Next
      For i = 3 To 4
          If pgb.Width >= 3000 Then
              lbTick(i).Top = lbTick(0).Top
              lbTick(i).Left = (lbTick(i - 3).Left + lbTick(i - 2).Left) / 2
              lbTick(i).Visible = True
            Else 'NOT PGB.WIDTH...
              lbTick(i).Visible = False
          End If
      Next i
    On Error GoTo 0
    Size pgb.Width, lbNum(0).Top + lbNum(0).Height

End Sub

Private Sub tmr_Timer()

  Dim F1 As Single, F2 As Single, v As Single, i As Long

    For i = 1 To ubDiff
        v = v + Diff(i)
        Diff(i - 1) = Diff(i)
    Next i
    i = 0
    RaiseEvent QueryDistance(i)
    On Error Resume Next
      Diff(ubDiff) = Abs(i - PrevValue) * 1000 / tmr.Interval
    On Error GoTo 0
    PrevValue = i
    v = (v + Diff(ubDiff)) / ubDiff
    If v Then
        Do
            F1 = IIf(Left$(lbNum(2), 1) = "2", 2.5, 2)
            F2 = IIf(Left$(lbNum(2), 1) = "5", 2.5, 2)
            Select Case v
              Case Is > pgb.Max * 0.9
                MaxSpeed = pgb.Max * F1
                If v <= pgb.Max Then
                    Exit Do '>---> Loop
                End If
              Case Is < pgb.Max / 4
                MaxSpeed = pgb.Max / F2
                If v > pgb.Max / F1 Then
                    Exit Do '>---> Loop
                End If
              Case Else
                Exit Do '>---> Loop
            End Select
        Loop
    End If
    pgb.Value = v

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        tmr.Interval = .ReadProperty("Interval", 1000)
        MaxSpeed = .ReadProperty("MaxSpeed", 1000)
        pgb.Value = .ReadProperty("CurrentSpeed", 0)
        Enabled = .ReadProperty("Enabled", True)
        ForeColor = .ReadProperty("ForeColor", &H80000012)
        BackColor = .ReadProperty("BackColor", &H8000000F)
        lbCapt.Caption = .ReadProperty("Caption", "")
    End With 'PROPBAG

End Sub

Private Sub UserControl_Show()

    If Virgin Then
        lbCapt.Caption = Parent.ActiveControl.Name
        UserControl_AmbientChanged ""
        Virgin = False
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Interval", tmr.Interval, 1000
        .WriteProperty "MaxSpeed", pgb.Max, 100
        .WriteProperty "CurrentSpeed", pgb.Value, 0
        .WriteProperty "Enabled", myEnabled, True
        .WriteProperty "ForeColor", lbNum(0).ForeColor, &H80000012
        .WriteProperty "BackColor", UserControl.BackColor, &H8000000F
        .WriteProperty "Caption", lbCapt.Caption, ""
    End With 'PROPBAG

End Sub

Private Function NextValid(ByVal Num As Long) As Long

  Dim i As Long, s As String

    i = 10 ^ (Int(Log(Num) / Log(10)))
    Do While i < Num
        s = Format$(i)
        If Left$(s, 1) = "2" Then
            i = i * 2.5
          Else 'NOT LEFT$(S,...
            i = i + i
        End If
    Loop
    NextValid = i

End Function

':) Ulli's VB Code Formatter V2.11.3 (11.04.2002 14:10:25) 8 + 258 = 266 Lines
