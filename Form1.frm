VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced DS2"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   1455
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt String"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt String"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "CipherText :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Password :"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Text :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim txt2 As String
Dim txt1 As String
Dim DS2 As New DS2
Dim st As String

txt1 = Text1.Text
txt2 = Text2.Text

If txt1 = "" Then GoTo 60
If txt2 = "" Then GoTo 80

GoTo 90

60 MsgBox "Enter text to be encrypted", vbOKOnly, "Advanced DS2"
   GoTo 100
80 MsgBox "Enter Password", vbOKOnly, "Advanced DS2"
   GoTo 100
   
90
   txt1 = encryptor(txt1, txt2)
   txt1 = encryptor(txt1, txt2)
   Text3.Text = DS2.EncryptString(txt1, txt2, True)

100

End Sub

Private Sub Command2_Click()
Dim txt2 As String
Dim txt1 As String
Dim txt3 As String
Dim DS2 As New DS2

txt1 = Text1.Text
txt2 = Text2.Text
txt3 = Text3.Text

If txt3 = "" Then GoTo 60
If txt2 = "" Then GoTo 80

GoTo 90

60 MsgBox "Enter Ciphertext", vbOKOnly, "Advanced DS2"
   GoTo 100
80 MsgBox "Enter Password", vbOKOnly, "Advanced DS2"
   GoTo 100

90 Text1.Text = DS2.DecryptString(txt3, txt2, True)
   Text1.Text = encryptor(Text1.Text, txt2)
   Text1.Text = encryptor(Text1.Text, txt2)
100
End Sub


Private Sub Command3_Click()
frmAbout.Show
End Sub

