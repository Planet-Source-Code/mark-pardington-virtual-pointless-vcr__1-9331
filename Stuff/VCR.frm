VERSION 5.00
Begin VB.Form VCR1 
   BackColor       =   &H00404040&
   Caption         =   "This is a VCR"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1200
      TabIndex        =   17
      ToolTipText     =   "These textboxs are where you enter in your details for what programmed timer recording you want"
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   16
      ToolTipText     =   "These textboxs are where you enter in your details for what programmed timer recording you want"
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer TimerRec 
      Left            =   120
      Top             =   720
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "These textboxs are where you enter in your details for what programmed timer recording you want"
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdChDn 
      Caption         =   "\/"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      ToolTipText     =   "These controls control what channel the video is currently set on"
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton CmdChUp 
      Caption         =   "/\"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      ToolTipText     =   "These controls control what channel the video is currently set on"
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "PARDY'S VCR"
      Height          =   795
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "This is the main video slot and as you can probably tell the video is currently inside"
      Top             =   360
      Width           =   5055
   End
   Begin VB.Timer TimerRecPlay 
      Left            =   6720
      Top             =   600
   End
   Begin VB.CommandButton Video 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      Height          =   495
      Left            =   1320
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "This is the main video slot and as you can probably tell the video is ejected"
      Top             =   480
      Width           =   4815
   End
   Begin VB.Timer TimerVideo 
      Left            =   6720
      Top             =   120
   End
   Begin VB.Timer TimerEject 
      Left            =   6240
      Top             =   120
   End
   Begin VB.Timer FadeLabels 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Channel 
      BackColor       =   &H80000012&
      Caption         =   "1"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      ToolTipText     =   "These controls control what channel the video is currently set on"
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label LblPos 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   135
      Index           =   2
      Left            =   240
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   135
      Index           =   1
      Left            =   240
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   375
      Left            =   240
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lbldisplay 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      ToolTipText     =   "This is the Main LCD screen display"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   240
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Eject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "These are the main controls of the VCR and perform all the different functions available"
      Top             =   1800
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      BorderWidth     =   5
      Height          =   495
      Left            =   1320
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "These are the main controls of the VCR and perform all the different functions available"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "FFWD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "These are the main controls of the VCR and perform all the different functions available"
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Set Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   4
      ToolTipText     =   "These are the main controls of the VCR and perform all the different functions available"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "These are the main controls of the VCR and perform all the different functions available"
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "These are the main controls of the VCR and perform all the different functions available"
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "REW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      ToolTipText     =   "These are the main controls of the VCR and perform all the different functions available"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      ToolTipText     =   "These are the main controls of the VCR and perform all the different functions available"
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "VCR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tmrec
starttime As String
endtime As String
Channel As String
End Type
Private Type Ctime
secs As Byte
mins As Byte
hours As Byte
End Type
Dim tms(100) As tmrec
Dim tmscount As Byte
Dim controly(50) As Integer
Dim byEjectCount As Byte
Dim ejected As Boolean
Dim curpos As Integer
Dim busy As Boolean
Dim currentfunc As String
Dim recordedtime As Ctime
Dim status As String
Dim addition As Integer
Dim aknowledged As Boolean


Public Sub FadeLabel(obj1 As Label, ByVal Index As Integer)
Dim x
    If obj1.ForeColor <> RGB(0, 0, 0) Then
        If controly(Index) > 0 Then
        controly(Index) = controly(Index) - 5
        obj1.ForeColor = RGB(controly(Index), 0, 0)
        End If
    End If
End Sub


Private Sub CmdChDn_Click()
If Val(Channel.Caption) > 1 And busy = False Then
Channel.Caption = Val(Channel.Caption) - 1
End If
End Sub

Private Sub CmdChUp_Click()

If Val(Channel.Caption) < 5 And busy = False Then
Channel.Caption = Val(Channel.Caption) + 1
End If
End Sub

Private Sub Form_Load()
curpos = 0
busy = False
Text1.Tag = 0
Text2.Tag = 0
Text3.Tag = 0
End Sub

Private Sub Label1_Click(Index As Integer)
Call VCRFunction(Label1(Index).Caption)
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If busy = False Then
Label1(Index).ForeColor = RGB(255, 0, 0)
controly(Index) = 255
lbldisplay.Caption = Label1(Index).Caption
End If
End Sub

Private Sub FadeLabels_Timer()
Dim x As Byte
    For x = 0 To 7
    Call FadeLabel(Label1(x), x)
    Next x
VCR1.Caption = Time
End Sub


Public Sub VCRFunction(ByVal functionname As String)
If ejected = False Then
TimerRecPlay.Enabled = True
    Select Case LCase$(functionname)
    Case "record"
        If busy = False Then
        busy = True
        TimerRecPlay.Interval = 1000
        TimerRecPlay.Enabled = True
        addition = 1
        status = "Recording"
        End If
    Case "play"
        If busy = False Then
        busy = True
        TimerRecPlay.Interval = 1000
        addition = 1
        status = "Playing"
        End If
    Case "rew"
     If busy = False Then
        busy = True
        TimerRecPlay.Interval = 1
        addition = -5
        status = "Rewinding"
        End If
    Case "stop"
        TimerRecPlay.Interval = 0
        busy = False
    Case "ffwd"
        If busy = False Then
        TimerRecPlay.Enabled = True
        busy = True
        TimerRecPlay.Interval = 1
        status = "Forwarding"
        addition = 5
        End If
    Case "program"
    Call Program
    
    Case "set time"
    Dim newtime As String
    Dim tmptime As String
    LblPos.Caption = ""
    tmptime = InputBox("Please enter the new hour value, the present hour is set to " & Hour(Now))
    If Len(tmptime) = 0 Then tmptime = "00"
    newtime = Format(tmptime, "00")
    tmptime = InputBox("Please enter the new minutes value, the present minutes is set to " & Minute(Now))
    If Len(tmptime) = 0 Then tmptime = "00"
    newtime = newtime & ":" & Format(tmptime, "00") & ":00"
    Time = newtime
    Case "eject"
        If busy = False Then Call Eject
    End Select
End If
End Sub

Public Sub Eject()
    If ejected <> True Then
    TimerRecPlay.Interval = 0
    byEjectCount = 0
    TimerEject.Interval = 10
    busy = True
    End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Text2.Text) > 59 Or Val(Text2.Text) < 0 Then
    MsgBox "Error you entered an invalid number"
    Text2.Text = ""
    Exit Sub
    Else
    Text3.SetFocus
    End If
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Text3.Text) > 59 Or Val(Text3.Text) < 0 Then
    MsgBox "Error you entered an invalid number"
    Text3.Text = ""
    Exit Sub
    Else
    aknowledged = True
    End If
End If
End Sub


Private Sub TimerRec_Timer()
Dim i, i2
For i = 1 To tmscount
If tms(i).starttime = Time Then
Call VCRFunction("record")
Channel.Caption = tms(i).Channel
End If
    If tms(i).endtime = Time Then
    Call VCRFunction("stop")
    For i2 = i To tmscount
        tms(i2).starttime = tms(i2 + 1).starttime
        tms(i2).endtime = tms(i2 + 1).endtime
        tms(i2).Channel = tms(i2 + 1).Channel
    Next i2
    tmscount = tmscount - 1
    End If
Next i
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Text1.Text) > 24 Or Val(Text1.Text) < 0 Then
    MsgBox "Error you entered an invalid number"
    Text1.Text = ""
    Exit Sub
    Else
    Text2.SetFocus
    End If
End If
End Sub

Private Sub TimerEject_Timer()
FadeLabels.Interval = 0
If Command1.Height >= 200 Then
Command1.Height = Command1.Height - 50
Else
Command1.Visible = False
byEjectCount = byEjectCount + 1
lbldisplay.Caption = "Ejecting"
    If byEjectCount = 10 Then
    TimerEject.Interval = 0
    FadeLabels.Interval = 10
    ejected = True
    busy = False
    Exit Sub
    End If
Video.Width = Video.Width + byEjectCount * 10
Video.Height = Video.Height + byEjectCount * 5
Video.Top = Video.Top - (byEjectCount * 5 / 2)
Video.Left = Video.Left - (byEjectCount * 10 / 2)

End If

End Sub

Private Sub TimerRecPlay_Timer()
Dim temporarypos
If curpos + addition < 10800 And curpos + addition >= 0 Then
curpos = curpos + addition
temporarypos = curpos
recordedtime.hours = Int(temporarypos / 3600)
temporarypos = temporarypos - (Int(temporarypos / 3600) * 3600)
recordedtime.mins = Int(temporarypos / 60)
temporarypos = temporarypos - (Int(temporarypos / 60) * 60)
recordedtime.secs = temporarypos
If recordedtime.secs > 60 Then
MsgBox "erm"
End If
lbldisplay.Caption = status
Call PositionCounter(recordedtime)
End If
If curpos + addition < 0 Then
TimerRecPlay.Interval = 0
Call Eject
End If
If curpos + addition >= 10800 Then
TimerRecPlay.Interval = 0
Call Eject
End If
End Sub

Private Sub TimerVideo_Timer()
byEjectCount = byEjectCount + 1
lbldisplay.Caption = "Loading"
    If byEjectCount > 9 Then
    Command1.Visible = True
    If Command1.Height <> 795 Then
    Command1.Height = Command1.Height + 50
    Exit Sub
    End If
    TimerVideo.Interval = 0
    ejected = False
    busy = False
    Exit Sub
    End If
Video.Width = Video.Width - byEjectCount * 10
Video.Height = Video.Height - byEjectCount * 5
Video.Top = Video.Top + (byEjectCount * 5 / 2)
Video.Left = Video.Left + (byEjectCount * 10 / 2)
End Sub

Private Sub Video_Click()
If busy = False Then
TimerRecPlay = 0
busy = True
byEjectCount = 0
    If ejected = True Then
    busy = True
    TimerVideo.Interval = 10
    End If
End If
End Sub

Private Sub PositionCounter(pos As Ctime)
Shape4(0).Width = Int((Shape2.Width / 59) * pos.secs)
Shape4(0).FillColor = RGB(0, 200, 0)
Shape4(0).FillStyle = 0
Shape4(1).Width = Int((Shape2.Width / 59) * pos.mins)
Shape4(1).FillColor = RGB(0, 200, 0)
Shape4(1).FillStyle = 0
Shape4(2).Width = Int((Shape2.Width / 3) * pos.hours)
Shape4(2).FillColor = RGB(0, 200, 0)
Shape4(2).FillStyle = 0
LblPos.Caption = Format(pos.hours, "00") & ":" & Format(pos.mins, "00") & ":" & Format(pos.secs, "00")
End Sub


Public Sub Program()
On Error GoTo exit1
LblPos.Caption = ""
tmscount = tmscount + 1
lbldisplay.Caption = "Start Time"
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text1.Text = Hour(Now)
Text2.Text = Minute(Now)
Text3.Text = Second(Now)
Text1.SetFocus
While aknowledged = False
lbldisplay.Caption = "Start Time"
DoEvents
Wend
aknowledged = False
tms(tmscount).starttime = Format(Text1.Text, "00") & ":" & Format(Text2.Text, "00") & ":" & Format(Text3.Text, "00")
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
lbldisplay.Caption = "Finish Time"
Text1.Text = Hour(Now)
Text2.Text = Minute(Now)
Text3.Text = Second(Now)
Text1.SetFocus
While aknowledged = False
lbldisplay.Caption = "Finish Time"
DoEvents
Wend
aknowledged = False
tms(tmscount).endtime = Format(Text1.Text, "00") & ":" & Format(Text2.Text, "00") & ":" & Format(Text3.Text, "00")
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.Visible = False
Text2.Visible = False
lbldisplay.Caption = "Channel"
Text3.Text = Channel.Caption
While aknowledged = False
lbldisplay.Caption = "Channel"
DoEvents
Wend
tms(tmscount).Channel = Text3.Text
MsgBox "OK"
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
TimerRec.Interval = 10
exit1:
Exit Sub
End Sub
