VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Progress Bar"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTestProgressBar 
      Caption         =   "Test Progress Bar"
      Height          =   540
      Left            =   1260
      TabIndex        =   0
      Top             =   1260
      Width           =   1695
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTestProgressBar_Click()

Dim ii As Integer
pgb.Show
For ii = 1 To 5000 ' ii = val , 5000 = maxval
    Call pgb.Progress(ii, 5000, "Testing the Progress Bar", &H8000&) ' Call the progressbar function
Next ii
Unload pgb

End Sub
