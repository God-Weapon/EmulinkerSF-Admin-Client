VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About EmulinkerSF Admin Client"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   8325
   Begin VB.TextBox txtAbout 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form7.frx":11F6
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Initialize()
    Me.Top = 0
    Me.Left = 3000
End Sub

Private Sub Form_Load()
    txtAbout.Text = "Version: EmulinkerSF Admin Client v" & clientVersion & vbCrLf
    txtAbout.Text = txtAbout.Text & "http://www.Supraclient.com" & vbCrLf
    txtAbout.Text = txtAbout.Text & "Compatibility: Emulinker Server 57.5+" & vbCrLf
    txtAbout.Text = txtAbout.Text & "http://www.emulinker.org" & vbCrLf
    txtAbout.Text = txtAbout.Text & "Protocol: Kaillera All Rights Reserved!" & vbCrLf
    txtAbout.Text = txtAbout.Text & "http://www.kaillera.com"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If allowUnload = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
    End If
End Sub

