VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   ScaleHeight     =   1950
   ScaleWidth      =   3600
   Begin VB.PictureBox picBar 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   0
      Width           =   3580
      Begin VB.PictureBox picExit 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   3300
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picMin 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   3060
         Picture         =   "Form1.frx":05BE
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picExit 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   3300
         Picture         =   "Form1.frx":0B7C
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   3
         Top             =   0
         Width           =   285
      End
      Begin VB.PictureBox picMin 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   3060
         Picture         =   "Form1.frx":113A
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   2
         Top             =   0
         Width           =   285
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MyApp"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   495
      End
   End
   Begin VB.Image imgStop 
      Height          =   450
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":16F8
      Top             =   1140
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Feel free to steal and redistribute... but please VOTE!!!"
      ForeColor       =   &H000080FF&
      Height          =   675
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   3015
   End
   Begin VB.Image imgStop 
      Height          =   450
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":1AC7
      Top             =   1140
      Width           =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GetPos As POINTAPI
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Dim MouseDown As Boolean
Dim FormX As Single
Dim FormY As Single
Private Type POINTAPI
    X As Long
    Y As Long
End Type


Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RELOAD
End Sub

Private Sub imgStop_Click(Index As Integer)
    End
End Sub
    
Private Sub imgStop_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgStop(1).Visible = False Then
        RELOAD
    End If
    imgStop(1).Visible = True
End Sub

Private Sub picbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = True
    FormX = X
    FormY = Y
End Sub

Private Sub picbar_mouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = 0
End Sub

Private Sub lbltitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RELOAD
    Call GetCursorPos(GetPos)
    If MouseDown = True Then
        Me.Top = (GetPos.Y * 15) - FormY
        Me.Left = (GetPos.X * 15) - FormX
    End If
End Sub

Private Sub lbltitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = True
    FormX = X
    FormY = Y
End Sub

Private Sub lbltitle_mouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = False
End Sub

Private Sub picbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RELOAD
    Call GetCursorPos(GetPos)
    If MouseDown = True Then
        Me.Top = (GetPos.Y * 15) - FormY
        Me.Left = (GetPos.X * 15) - FormX
    End If
End Sub

Sub RELOAD()
    imgStop(1).Visible = False
    picMin(1).Visible = False
    picExit(1).Visible = False
End Sub

Private Sub picExit_Click(Index As Integer)
    End
End Sub

Private Sub picExit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picExit(1).Visible = False Then
        RELOAD
    End If
    picExit(1).Visible = True
End Sub

Private Sub picMin_Click(Index As Integer)
    Me.WindowState = vbMinimized
End Sub

Private Sub picMin_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picMin(1).Visible = False Then
        RELOAD
    End If
    picMin(1).Visible = True
End Sub
