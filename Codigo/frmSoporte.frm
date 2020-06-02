VERSION 5.00
Begin VB.Form frmSoporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soporte"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEmail 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtPersonaje 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtRespuesta 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   480
      MaxLength       =   300
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmSoporte.frx":0000
      Top             =   4200
      Width           =   7695
   End
   Begin VB.CommandButton btnResponder 
      Caption         =   "RESPONDER"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox txtMensaje 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmSoporte.frx":000A
      Top             =   3000
      Width           =   7695
   End
   Begin VB.ListBox lstSoportes 
      Height          =   1425
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   7695
   End
   Begin VB.TextBox txtSoportesSinRespuesta 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblEmailPj 
      Caption         =   "eMail:"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblPersonaje 
      Alignment       =   1  'Right Justify
      Caption         =   "Personaje:"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblRespChars 
      Caption         =   "300/300"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblSoportesSinResponder 
      Caption         =   "Soportes sin respuesta:"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ListaSoportes As String
Dim ArraySoportes(100) As String
Dim UserIndexReclamo As Integer
Dim IdReclamo As Integer

Private Sub btnResponder_Click()
    Call SendData("RGM" & IdReclamo & "|" & UserIndexReclamo & "|" & txtPersonaje.Text & "|" & txtEmail.Text & "|" & txtRespuesta.Text)
    Unload frmSoporte
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim CantSinResponder As Integer
    
    CantSinResponder = 0
    
    For i = 0 To 99
        ArraySoportes(i) = ReadField(i + 1, ListaSoportes, 59)
    Next i
    
    For i = 0 To 99
        If ArraySoportes(i) <> "" Then
            CantSinResponder = CantSinResponder + 1
            lstSoportes.AddItem ReadField(1, ArraySoportes(i), 124) & " - (" & ReadField(2, ArraySoportes(i), 124) & ") - " & ReadField(4, ArraySoportes(i), 124) & " - " & ReadField(7, ArraySoportes(i), 124)
        End If
    Next i
    
    txtSoportesSinRespuesta = CantSinResponder
    
End Sub

Private Sub lstSoportes_Click()
    If lstSoportes.ListIndex >= 0 Then
        IdReclamo = ReadField(1, ArraySoportes(lstSoportes.ListIndex), 124)
        UserIndexReclamo = ReadField(4, ArraySoportes(lstSoportes.ListIndex), 124)
        
        txtRespuesta.Text = ""
        txtRespuesta.Enabled = True
        btnResponder.Enabled = True
        
        txtPersonaje.Text = ReadField(5, ArraySoportes(lstSoportes.ListIndex), 124)
        txtEmail.Text = ReadField(6, ArraySoportes(lstSoportes.ListIndex), 124)
        
        txtMensaje.Text = ReadField(9, ArraySoportes(lstSoportes.ListIndex), 124)
    End If
    
End Sub

Private Sub txtRespuesta_Change()
    lblRespChars.Caption = (300 - Len(txtRespuesta.Text)) & "/300"
End Sub
