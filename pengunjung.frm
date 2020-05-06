VERSION 5.00
Begin VB.Form pengunjung 
   BackColor       =   &H8000000E&
   Caption         =   "DATA PEGUNJUNG"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "OK"
      Height          =   855
      Left            =   1320
      MaskColor       =   &H00808000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   6855
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   2760
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   6855
   End
End
Attribute VB_Name = "pengunjung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text4.Text = Format(Now, "dd MMMM yyyy")
Text5.Text = Format(Now, "hh : mm : ss")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If ascii = 13 Then
MsgBox "FORM USERNAME ANDA MASIH KOSONG !", vbCritical, "Perhatian"
End If
End Sub
