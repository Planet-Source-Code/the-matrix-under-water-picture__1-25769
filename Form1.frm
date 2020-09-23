VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make Picture Under Water // By The Matrix "" Qais Ghalib ""  "
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   0
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4695
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   6525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Qais Ghalib
'the Matrix
'part ss114
'Qais60@hotmail.com

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDC As Long, ByVal rc As Long, ByVal yrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Type Clip
x As Single
y As Single
Set As Single
offsety As Single
End Type
Dim Verb() As Clip, Sine1() As Single, Sine2() As Single, i As Integer, ii As Integer
Const Siz = 16

Private Sub underwater()
Do: DoEvents
For i = 0 To Picture1.ScaleWidth / Siz
For ii = 0 To Picture1.ScaleHeight / Siz
Verb(i, ii).x = i * Siz
Verb(i, ii).y = ii * Siz
Verb(i, ii).Set = Verb(i, ii).Set + 0.5 * Sin(Timer * i)
Verb(i, ii).offsety = Verb(i, ii).offsety + 0.5 * Sin(Timer * ii)
BitBlt Picture1.hDC, Verb(i, ii).x + Verb(i, ii).Set, Verb(i, ii).y + Verb(i, ii).offsety, Siz, Siz, Picture2.hDC, i * Siz, ii * Siz, SRCCOPY
Next
Next
Loop
End Sub

Private Sub Command1_Click()
underwater
End Sub

Private Sub Form_Activate()
underwater
End Sub

Private Sub Form_Load()
ReDim Verb(Picture1.ScaleWidth / Siz, Picture1.ScaleHeight / Siz)
End Sub

