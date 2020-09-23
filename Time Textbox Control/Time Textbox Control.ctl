VERSION 5.00
Begin VB.UserControl Time 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1125
   KeyPreview      =   -1  'True
   ScaleHeight     =   315
   ScaleWidth      =   1125
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
      Begin VB.TextBox txtAMPM 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   625
         TabIndex        =   2
         Text            =   "AM"
         Top             =   15
         Width           =   345
      End
      Begin VB.TextBox txtMinutes 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   390
         TabIndex        =   1
         Text            =   "01"
         Top             =   15
         Width           =   225
      End
      Begin VB.TextBox txtColon 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   ":"
         Top             =   15
         Width           =   105
      End
      Begin VB.TextBox txtHours 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Text            =   "01"
         Top             =   15
         Width           =   225
      End
   End
End
Attribute VB_Name = "Time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub txtAMPM_GotFocus()
txtAMPM.SelStart = 0
txtAMPM.SelLength = Len(txtAMPM.Text)
End Sub

Private Sub txtHours_GotFocus()
txtHours.SelStart = 0
txtHours.SelLength = Len(txtHours.Text)
End Sub

Private Sub txtMinutes_GotFocus()
txtMinutes.SelStart = 0
txtMinutes.SelLength = Len(txtMinutes.Text)
End Sub

Private Sub txtHours_LostFocus()
txtHours = Right(txtHours, 2)
If Left(txtHours.Text, 1) <> 0 And txtHours.Text <= 9 Then txtHours.Text = "0" & txtHours.Text
If txtHours.Text > 12 Then txtHours.Text = "12"
End Sub

Private Sub txtMinutes_LostFocus()
txtMinutes = Right(txtMinutes, 2)
If Left(txtMinutes.Text, 1) <> 0 And txtMinutes.Text <= 9 Then txtMinutes.Text = "0" & txtMinutes.Text
If txtMinutes.Text > 59 Then txtMinutes.Text = "59"
End Sub

Private Sub txtHours_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 38
        If txtHours.Text >= 12 Then txtHours = 12 Else txtHours.Text = txtHours.Text + 1
    Case 40
        If txtHours.Text > 12 Then txtHours = 12 Else txtHours.Text = txtHours.Text - 1
    Case 37
    Case 39
End Select

If txtHours.Text < 1 Then txtHours.Text = "1"

End Sub

Private Sub txtMinutes_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 38
        If txtMinutes.Text >= 59 Then txtMinutes = 59 Else txtMinutes.Text = txtMinutes.Text + 1
    Case 40
        If txtMinutes.Text > 59 Then txtHours = 59 Else txtMinutes.Text = txtMinutes.Text - 1
    Case 37
    Case 39
End Select

If txtMinutes.Text < 1 Then txtMinutes.Text = "1"

End Sub

Private Sub txtAMPM_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 38
        txtAMPM = "AM"
    Case 40
        txtAMPM = "PM"
    Case Else
        KeyCode = 0
End Select

txtAMPM.SelStart = 0
txtAMPM.SelLength = Len(txtAMPM.Text)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = 0
End Sub

