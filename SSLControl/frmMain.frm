VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SSL Winsock With Proxy Support"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   1440
      TabIndex        =   2
      Top             =   4200
      Width           =   3735
      Begin VB.CheckBox Check1 
         Caption         =   "Use Proxy Server"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "80.66.177.4:80 "
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin sslControl.SSLSocket SSLSocket1 
      Left            =   3120
      Top             =   2640
      _extentx        =   741
      _extenty        =   741
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = 1 Then
    SSLSocket1.CloseSSL
    SSLSocket1.ConnectSSL "paypal.com", , Text2
Else
    SSLSocket1.CloseSSL
    SSLSocket1.ConnectSSL "paypal.com"
End If
End Sub

Private Sub sslsocket1_SSLConnect()
    SSLSocket1.SendSSL "GET / HTTP/1.0" & vbCrLf & vbCrLf
End Sub

Private Sub sslsocket1_SSLData(sData As String)
    Text1 = sData
End Sub

Private Sub SSLSocket1_SSLProxyError()
    MsgBox "Your proxy is no good"
End Sub
