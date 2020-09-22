VERSION 5.00
Begin VB.Form pgb 
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimate 
      Interval        =   100
      Left            =   4680
      Top             =   345
   End
   Begin VB.PictureBox Picture1 
      Height          =   250
      Left            =   120
      ScaleHeight     =   195
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   1050
      Width           =   5580
      Begin VB.CheckBox chkPrg 
         BackColor       =   &H00008000&
         Height          =   200
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   105
      Width           =   5535
   End
   Begin VB.Image imgGo 
      Height          =   240
      Index           =   0
      Left            =   360
      Picture         =   "pgb.frx":0000
      Top             =   465
      Width           =   240
   End
   Begin VB.Image imgGo 
      Height          =   240
      Index           =   1
      Left            =   2760
      Picture         =   "pgb.frx":03B6
      Top             =   465
      Width           =   240
   End
   Begin VB.Image imgMiddle 
      Height          =   480
      Left            =   2520
      Picture         =   "pgb.frx":075A
      Top             =   345
      Width           =   480
   End
   Begin VB.Image imgEnd 
      Height          =   480
      Left            =   5160
      Picture         =   "pgb.frx":2EFC
      Top             =   345
      Width           =   480
   End
   Begin VB.Image imgStart 
      Height          =   480
      Left            =   120
      Picture         =   "pgb.frx":3206
      Top             =   345
      Width           =   480
   End
   Begin VB.Label lblPerc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0% Completed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1365
      Width           =   5535
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5535
   End
End
Attribute VB_Name = "pgb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Author : Renier Barnard (renier_barnard@santam.co.za)
'
' Date    : May 2000
'
' Description :
' This code will demonstrate how to make a simple but nice
' looking progress bar. It could be more simple (Using the line command)
' but this looks better... way better. The Status bar also changes colour as it progresses.
' It will also calculate the time remaining and display it
' In addition to this , it animates some icons to keep the form "busy looking"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const FLAGS = 1
Const HWND_TOPMOST = -1
Dim Aindex As Integer
Dim LastPos As Long
Dim lLastTime As Double
Dim tLastTime


Public Function Progress(Value, MaxValue, Optional HeaderX As String, Optional color As ColorConstants)
'' This is the actual progress bar function.

DoEvents
Dim Perc
Dim bb As Integer
Dim lTime As Double
Dim lTimeDiff As Double
Dim lTimeLeft As Double
Dim lTotalTime As Double
'Me.Show

'Get a color to do it in
If color = 0 Then color = &H8000&

'Display the header , if any was returned
If HeaderX <> "" Then
    lblHeader = HeaderX
Else
    lblHeader = "Busy Processing...Please wait"
End If

'Now work out the percentage (0-100) of where we currently are
Perc = (Value / MaxValue) * 100
If Perc < 0 Then Perc = 0
If Perc > 100 Then Perc = 100
Perc = Int(Perc)

'Do the time remaining calculation
If (Perc Mod 10) = 0 Or Perc = 0 Then ' Every 10 percent
        lTimeDiff = lTime - lLastTime
        lTime = Time - tLastTime
        If Perc = 0 Or Perc < 0 Then
            lTotalTime = ((100 / 1) * 2) * lTime
            lTimeLeft = (((100 / 1) * 2) * lTime) - (((100 / 100) * 2) * lTime)
        Else
            lTotalTime = ((100 / Perc) * 2) * lTime
            lTimeLeft = (((100 / Perc) * 2) * lTime) - (((100 / 100) * 2) * lTime)
        End If
        lblTime = "Time Remaining : " & Format((lTimeLeft), "hh:mm:ss")
        lblTime.ForeColor = color
End If
DoEvents


lblPerc.ForeColor = color
lblHeader.ForeColor = color
chkPrg.BackColor = color
DoEvents
lblPerc.Caption = Int(Perc) & "% Completed" 'Just the Label Display


chkPrg.BackColor = RGB(0, Perc * 2.5, 255 - Perc * 2.5)

chkPrg.Width = Int(Perc)

DoEvents

End Function



Private Sub Form_Load()

tLastTime = Time

'Const FLAGS = 1
'Const HWND_TOPMOST = -1
Aindex = 0
LastPos = 720

'Me.Width = 5910
'Me.Height = 1545

Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2

'Sets form on always on top.
Dim Success As Integer
'Success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
                                                ' Change the "0's" above to position the window.

Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
DoEvents

End Sub


Private Sub Form_Unload(Cancel As Integer)
DoEvents
End Sub

Private Sub tmrAnimate_Timer()
'This funtion will animate a couple of icons , just to show that something is busy hapening

DoEvents
LastPos = LastPos + 1

If LastPos > 2680 And LastPos < 3250 Then
    LastPos = 3160
    Aindex = 1
Else
    If LastPos > 5360 Then
        LastPos = 720
        Aindex = 0
    Else
        
    End If
End If

If Aindex = 1 Then
    imgGo(1).Visible = True
    imgGo(0).Visible = False
Else
    imgGo(1).Visible = False
    imgGo(0).Visible = True
End If

LastPos = LastPos + 200
imgGo(Aindex).Left = LastPos
DoEvents

End Sub


