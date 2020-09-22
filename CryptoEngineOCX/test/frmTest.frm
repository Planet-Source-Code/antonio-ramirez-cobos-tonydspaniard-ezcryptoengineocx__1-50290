VERSION 5.00
Object = "*\A..\CryptoEngine.vbp"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5055
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   900
      Width           =   3825
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5025
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   240
      Width           =   3915
   End
   Begin vbCryptoEngine.CryptoEngine CryptoEngine1 
      Left            =   3315
      Top             =   225
      _ExtentX        =   1614
      _ExtentY        =   1614
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Twofish"
      Height          =   375
      Index           =   6
      Left            =   2400
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "TEA"
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Skipjack"
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Rijndael"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Gost"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DES"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Blowfish"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDecrypt_Click()
On Error Resume Next
    'CryptoEngine1.DecryptFile "F:\vb\cryptoengineocx\trucos2.txt", "F:\vb\cryptoengineocx\trucos2.txt", True
    Text2 = CryptoEngine1.DecryptString(Text2, "Robert", True)
    
End Sub

Private Sub cmdEncrypt_Click()
On Error Resume Next
    'CryptoEngine1.EncryptFile "F:\vb\cryptoengineocx\trucos.txt", "F:\vb\cryptoengineocx\trucos2.txt", True
    Text2 = CryptoEngine1.EncryptString(Text1, "Robert", True)
End Sub

Private Sub CryptoEngine1_Error(Number As Long, Source As String, Description As String)
    Debug.Print Number; Source; Description
End Sub

Private Sub CryptoEngine1_Process(percent As Long)
    Me.Caption = percent & "%"
End Sub

Private Sub CryptoEngine1_StatusChanged(lstatus As Long)
    Debug.Print lstatus
End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
     CryptoEngine1.CryptAlgorithm = Index
End Sub
