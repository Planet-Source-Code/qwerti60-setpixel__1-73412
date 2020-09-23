VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   4665
   ClientTop       =   3645
   ClientWidth     =   5040
   DrawWidth       =   10
   FillColor       =   &H00004000&
   ForeColor       =   &H00004040&
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form3.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   2280
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog Com1 
      Left            =   2160
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1080
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   1200
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 27 Then Form_Unload (0) ' Press Enter or ESC to end the programme
If KeyCode = 32 Then Timer1_Timer 'Press spacebar to start a new drawing
If KeyCode = Asc("1") Then
Me.Cls 'when the key 1 is pressed, the screen is cleared
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("P") Or KeyAscii = Asc("p") Then Pau = Not Pau 'P/p to pause/play the drawing
If KeyAscii = 32 Then Timer1_Timer 'Press spacebar to start a new drawing
If KeyAscii = Asc("1") Then
Me.Cls 'when the key 1 is pressed, the screen is cleared
End If
End Sub

Private Sub Form_Load()
Randomize Timer
Randomize Rnd
Form3.AutoRedraw = True
hDc1 = GetWindowDC(GetDesktopWindow)
BitBlt Form3.hdc, 0, 0, Form3.Width, Form3.Height, hDc1, 0, 0, &HCC0020
'copies the screen to the form
Form3.AutoRedraw = False

'Com1.ShowSave
'ShowCursor 0
'Random numbers for the math functions
RR1(1) = 63496252454.81
RR1(2) = 63488783024.57: N2(2) = 918046539776#
RR1(3) = 63543968745.75: N2(3) = 7977.56
RR1(4) = 63496612813.63: N2(4) = 675
RR1(5) = 63496612813.63: N2(5) = 127
RR1(6) = 63595311370.3: N2(6) = 801.051
RR1(7) = 63494682649.15: N2(7) = 119

RR1(8) = 63496977370.38
Pau = False
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_Unload (0) 'when the mouse is pressed the programme ends
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_Unload (0) 'when the mouse is pressed the programme ends
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowCursor 1 ' To show the mouse cursor in VB6
MsgBox "The programme was written by Roman Braverman", , "Pictures"
End
End Sub

Private Sub Timer1_Timer() ' To draw the different pictures
Randomize Timer
Randomize Rnd
D1
End Sub

Private Sub Timer2_Timer() ' to stop the drawing
DoEvents
If (XX <= X) Or (X >= XX) Or (XX = X) Or (XX = (Form3.ScaleWidth / 2)) Or (X = (Form3.ScaleWidth / 2)) Then
SLEEP1
Timer1_Timer
End If
End Sub

Private Sub Timer3_Timer()
'a loop to pause the drawing
DoEvents
If Pau = True Then
Do
DoEvents
Loop While Pau = True
End If
End Sub
