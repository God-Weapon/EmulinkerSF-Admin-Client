VERSION 5.00
Begin VB.Form frmTrivia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trivia Bot Commands"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   4230
   Begin VB.CommandButton btnWin 
      Caption         =   "Show Winner"
      Height          =   315
      Left            =   1800
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton btnTop3 
      Caption         =   "Top 3 Scores"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset Bot"
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton BtnTime 
      Caption         =   "Question Delay Time"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save Scores"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtSec 
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Text            =   "30"
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton btnPause 
      Caption         =   "Pause Bot"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton btnOn 
      Caption         =   "Turn Bot ON"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton btnOff 
      Caption         =   "Turn Bot OFF"
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton btnResume 
      Caption         =   "Resume Bot"
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmTrivia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim g As Long


Private Sub btnOn_Click()
    Call globalChatRequest("/triviaon")
End Sub

Private Sub btnOff_Click()
    Call globalChatRequest("/triviaoff")
End Sub

Private Sub btnPause_Click()
    Call globalChatRequest("/triviapause")
End Sub

Private Sub btnResume_Click()
    Call globalChatRequest("/triviaresume")
End Sub

Private Sub btnSave_Click()
    Call globalChatRequest("/triviasave")
End Sub

Private Sub btnReset_Click()
    Call globalChatRequest("/triviareset")
End Sub

Private Sub btnTop3_Click()
    Call globalChatRequest("/triviascores")
End Sub

Private Sub btnWin_Click()
    Call globalChatRequest("/triviawin")
End Sub


Private Sub btnTime_Click()
    Call globalChatRequest("/triviatime " & txtSec.Text)
End Sub

Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 4900
End Sub

Private Sub Form_Load()
    g = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
    End If
End Sub
