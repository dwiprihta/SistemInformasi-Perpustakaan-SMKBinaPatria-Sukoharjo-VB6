VERSION 5.00
Begin VB.Form index 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   Caption         =   "PERPUSTAKAAN SMK BINA PATRIA 2 SUKOHARJO"
   ClientHeight    =   10755
   ClientLeft      =   75
   ClientTop       =   720
   ClientWidth     =   20370
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "index.frx":0000
   ScaleHeight     =   10755
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--:--:--"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   10680
      TabIndex        =   2
      Top             =   7440
      Width           =   2895
   End
   Begin VB.Label Label88 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--/--/----"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   7440
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   15120
      TabIndex        =   0
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Menu MASTER 
      Caption         =   "MASTER"
      Begin VB.Menu DABUK 
         Caption         =   "DATA BUKU"
      End
      Begin VB.Menu DATAANG 
         Caption         =   "DATA ANGGOTA"
         Begin VB.Menu siswa 
            Caption         =   "SISWA"
         End
         Begin VB.Menu dosen 
            Caption         =   "DOSEN & KARYAWAN"
         End
      End
      Begin VB.Menu DATRANS 
         Caption         =   "DATA TRANSAKSI SELESAI"
      End
   End
   Begin VB.Menu TRANS 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu pin 
         Caption         =   "PEMINJAMAN"
      End
      Begin VB.Menu PENG 
         Caption         =   "PENGEMBALIAN"
      End
   End
   Begin VB.Menu OUT 
      Caption         =   "LOG-OUT"
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'FORM HALAMAN UTAMA APLIKASI
'MENAMPILKAN MENU KE SELURUH FORM
'by INDRI DWI S
'======================================================================

'TAMPILKAN FORM DATA KEMBALI BUKU
Private Sub DAPENG_Click()
kembali.Show
End Sub

'TAMPILKAN FORM DATA BUKU
Private Sub DABUK_Click()
data_buku.Show
End Sub

'TAMPILKAN FORM DATA ANGGOTA
Private Sub DATAANG_Click()

End Sub

'TAMPILKAN FORM DATA TRANSAKSI
Private Sub DATRANS_Click()
trans_selesai.Show
End Sub

Private Sub dosen_Click()
anggota_guru.Show
End Sub

'TAMPILKAN WAKTU
Private Sub Label2_Click()
Label2.Caption = Time
End Sub

'PERTANYAAN SAAT AKAN KELUAR
Private Sub OUT_Click()
If MsgBox("Apakah Anda yakin ingin keluar ?", vbYesNo + vbDefaultButton2 + vbQuestion, "VB 6.0 WARNING !") = vbYes Then
End
End If
End Sub

'TAMPILKAN FORM TRANSAKSI PEMINJAMAN
Private Sub PEMIN_Click()
trans_pinjam.Show
End Sub

'TAMPILKAN FORM PINJAM BUKU
Private Sub PENG_Click()
trans_kembali.Show
End Sub

'TAMPILKAN FORM TRANSAKSI PEMINJAMAN
Private Sub pin_Click()
trans_pinjam.Show
End Sub

Private Sub siswa_Click()
data_anggota.Show
End Sub

'TAMPILKAN WAKTU
Private Sub Timer1_Timer()
Label77.Caption = Format(Now, "hh : mm : ss")
'Label88.Caption = Format(Now, "dd MMMM yyyy")
Label88.Caption = Format(Now, "dd MMMM yyyy")
   'Label2.Caption = Time
End Sub




