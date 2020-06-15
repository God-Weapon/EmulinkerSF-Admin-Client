VERSION 5.00
Begin VB.Form frmAscii 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASCII ART"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   7950
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmAscii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub btnSend_Click()
    Dim str() As String
    Dim i As Long
    Dim w As Long
    
    str = Split(Text1.Text, vbCrLf)
    For i = 0 To UBound(str)
        Call splitAnnounce(str(i))
        w = GetTickCount
        Do Until GetTickCount - w >= 20
            DoEvents
        Loop
    Next i
End Sub

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
    End If
End Sub
