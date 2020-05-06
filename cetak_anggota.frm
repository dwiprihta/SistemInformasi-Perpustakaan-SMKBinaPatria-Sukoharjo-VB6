VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form cetak_anggota 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7065
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   840
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   5040
      ScaleHeight     =   2115
      ScaleWidth      =   1875
      TabIndex        =   9
      Top             =   240
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   1200
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   2040
      Width           =   3750
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1200
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   3630
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "CETAK"
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin Crystal.CrystalReport crLAP6 
      Left            =   240
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "cetak_anggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FORM CARI CETAK DATA ANGGOTA (PERORANGAN)
'MENAMPILKAN DATA ANGGOTA YANG AKAN DICETAK SECARA INDIVIDUAL
'by INDRI DWI S
'======================================================================

'jika tombol 1 ditekan
Private Sub Command1_Click()
' laoran 6 dan dimasukan kedalam variabel xx
xx = "\LAP6.rpt"
cc = "*"
With crLAP6
    .ReportFileName = App.Path & xx
    .WindowState = crptMaximized
    'cetak data kartu berdasarkan data dibawah ini
    .Formulas(0) = "NAP='" & Text1.Text & "'"
    .Formulas(1) = "Nama='" & TEXT3.Text & "'"
    .Formulas(2) = "Kelas='" & TEXT3.Text & "'"
    .Formulas(3) = "Alamat='" & Text4.Text & "'"
    .Formulas(4) = "image= '" & Image1.Picture & "'"
    .RetrieveDataFiles
    .Action = 1
    End With
End Sub

Private Sub Command2_Click()
'jika user batal mencetak data
Unload Me
Form4.Show
End Sub

