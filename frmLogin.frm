VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   6225
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3677.936
   ScaleMode       =   0  'User
   ScaleWidth      =   4436.519
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text1 
      Height          =   585
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   4005
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "LOGIN"
      Default         =   -1  'True
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1500
   End
   Begin VB.TextBox text2 
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4320
      Width           =   4005
   End
   Begin VB.Image Image1 
      Height          =   2265
      Left            =   1200
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "USERNAME"
      Height          =   270
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PASSWORD"
      Height          =   270
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'FORM LOGIN
'MENAMPILKAN FORM LOGIN
'by INDRI DWI S
'======================================================================

Private Sub cmdOK_Click()
'panggil modul koneksi
Call Koneksi
'cek jika form masih kosong
If Text1.Text = "" Then
MsgBox "FORM USERNAME ANDA MASIH KOSONG !", vbCritical, "Perhatian"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "FORM PASSWORD ANDA MASIH KOSONG !!!", vbCritical, "Perhatian"
Text2.SetFocus
Else

'cari data login di database admin
query = "select * from login where username='" & Text1.Text & "' and password='" & Text2.Text & "'"
RS.Open (query), conn
    If RS.EOF Then
    'tampilkan notif jika username atau password salah
    MsgBox "USERNAME ATAU PASSWORD ANDA SALAH !", vbExclamation, "Gagal !"
    'bersihkan inputan form
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Else
    
    'jika berhasil login masuk ke menu admin
    MsgBox "ANDA BERHASIL LOGIN !", vbInformation, "LOGIN SUKSES !"
    index.Show
    'tutup form login
    Unload Me
    End If
End If
End Sub

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub


