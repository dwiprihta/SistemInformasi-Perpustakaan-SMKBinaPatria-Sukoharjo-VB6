VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form anggota_guru 
   BackColor       =   &H8000000E&
   Caption         =   "ANGGOTA PERPUSTAKAAN (GURU DAN KARYAWAN)"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      DataField       =   "Foto"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2520
      TabIndex        =   31
      Text            =   "Text8"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6240
      Width           =   1410
   End
   Begin VB.CommandButton Commandcari 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16440
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   9255
      Begin VB.ComboBox TEXT3 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   32
         Top             =   4440
         Width           =   5175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "PILIH FOTO"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         TabIndex        =   18
         Top             =   2640
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         Height          =   2295
         Left            =   4320
         ScaleHeight     =   2235
         ScaleWidth      =   1755
         TabIndex        =   17
         Top             =   240
         Width           =   1815
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            Height          =   2535
            Left            =   0
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1935
         End
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   11
         Top             =   3240
         Width           =   5175
      End
      Begin VB.TextBox Text2 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   10
         Top             =   3840
         Width           =   5175
      End
      Begin VB.ComboBox Combo1 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   9
         Top             =   5040
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Anggota       "
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
         Left            =   600
         TabIndex        =   15
         Top             =   3240
         Width           =   5535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama                              "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   3960
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Status            "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin              "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   5160
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   5775
      Left            =   10200
      TabIndex        =   1
      Top             =   240
      Width           =   9735
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "Tambah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3360
         Width           =   2490
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H8000000D&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3960
         Width           =   2490
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H8000000D&
         Caption         =   "Hapus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3960
         Width           =   2490
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000D&
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3360
         Width           =   2490
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H8000000D&
         Caption         =   "Cetak Perorangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5160
         Width           =   5265
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H8000000D&
         Caption         =   "Cetak Semua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4680
         Width           =   5265
      End
      Begin VB.TextBox Text7 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         TabIndex        =   24
         Top             =   2760
         Width           =   5175
      End
      Begin VB.TextBox Text4 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         TabIndex        =   23
         Top             =   360
         Width           =   5175
      End
      Begin VB.TextBox Text5 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         TabIndex        =   22
         Top             =   960
         Width           =   5175
      End
      Begin VB.TextBox Text6 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4080
         TabIndex        =   21
         Top             =   1560
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   2160
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   59834369
         CurrentDate     =   43534
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Foto              "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   2880
         Width           =   4575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Mulai            "
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
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Sekarang      "
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
         Left            =   480
         TabIndex        =   5
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat dan Tanggal Lahir "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "No Induk Pegawai      "
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
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   12840
      TabIndex        =   0
      Top             =   6240
      Width           =   3255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "anggota_guru.frx":0000
      Height          =   3375
      Left            =   600
      TabIndex        =   16
      Top             =   6960
      Width           =   19335
      _ExtentX        =   34105
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   -2147483644
      HeadLines       =   2
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport crLAP5 
      Left            =   1440
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\SOFTWARE PERPUSTAKAAN\perpus.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\SOFTWARE PERPUSTAKAAN\perpus.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ANGGOTAPERPUSTAKAAN"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "anggota_guru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FORM DATA ANGGOTA
'MENAMPILKAN DATA ANGGOTA DAN OPERASI (SIMPAN, HAPUS, UBAH, DAN CETAK)
'by INDRI DWI S
'======================================================================

'kode anggota otomatis
Sub KodeOtomatis()
Call Koneksi
RS.Open ("select * from ANGGOTAPERPUSTAKAAN Where NomorAnggota In(Select Max(NomorAnggota)From ANGGOTAPERPUSTAKAAN)Order By NomorAnggota Desc"), conn
RS.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RS
        If .EOF Then
            Urutan = "AGT" + "001"
            Text1 = Urutan
        Else
            Hitung = Right(!NomorAnggota, 3) + 1
            Urutan = "AGT" + Right("000" & Hitung, 3)
        End If
        Text1 = Urutan
    End With
End Sub

'tampilka data dalam tabel
Sub query()
Adodc1.Recordset.Filter = "kelas like '%r%'"
'set lebar kolom pada tabel
With DataGrid1
.Columns(0).Width = 2000
.Columns(1).Width = 3000
.Columns(2).Width = 1000
.Columns(3).Width = 1500
.Columns(4).Width = 2000
.Columns(5).Width = 3500
.Columns(6).Width = 3500
.Columns(7).Width = 1300
.Columns(8).Width = 1800

.Columns(0).Caption = "NO ANGGOTA"
.Columns(1).Caption = "NAMA ANGGOTA"
.Columns(2).Caption = "SATUS"
.Columns(3).Caption = "JENIS KELAMIN"
.Columns(4).Caption = "NO INDUK PEGAWAI"
.Columns(5).Caption = "TEMPAT, TGL LHR"
.Columns(6).Caption = "ALAMAT"
.Columns(7).Caption = "MULAI ANGGOTA"
.Columns(8).Caption = "FOTO"
End With
End Sub

'perintah otomatis yang dijalankan saat form data anggota dibuka
Private Sub Form_Load()
'seting waktu pada datagrid
With DataGrid1
.Columns(7).NumberFormat = "dd MMMM yy"
End With

Call KodeOtomatis
Call query

'isikan nilai dari combo 1 (jenis kelamin)
Combo1.AddItem "Laki-laki"
Combo1.AddItem "Perempuan"
TEXT3.AddItem "Guru"
TEXT3.AddItem "Karyawan"
End Sub

'membuat variabel untuk merefresh data (dipanggil pada tombol tambah, edit, hapus)
Sub tblrfrsh()
Adodc1.Refresh
With DataGrid1
.Columns(0).Width = 2000
.Columns(1).Width = 3000
.Columns(2).Width = 1000
.Columns(3).Width = 1500
.Columns(4).Width = 2000
.Columns(5).Width = 3500
.Columns(6).Width = 3500
.Columns(7).Width = 1300
.Columns(8).Width = 1800
End With
End Sub

'membuat variabel untuk membersihkan data pada form (dipanggil pada tombol tambah, edit, hapus)
Sub bersih()
'Text1.Text = ""
Text2.Text = ""
TEXT3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
DTPicker1.Value = Now
End Sub

'membuat variabel untuk membuat form menjadi hidup (dipanggil pada tombol tambah)
Sub enabel()
'Text1.Enabled = True
Text2.Enabled = True
TEXT3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Combo1.Enabled = True
Combo1.Enabled = True
Command7.Enabled = True
Text2.SetFocus
End Sub

'jika tombol tambah di tekan
Private Sub Command1_Click()
'memanggil variabel bersih
Command2.Enabled = True
Call bersih
'memeanggil variabel untuk menghidupkan form
Call enabel
Call KodeOtomatis
End Sub

'SIMPAN
'jika tombol simpan diklik
'script untuk menyimpan data anggota
Private Sub Command2_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text1 = "" Or Text2 = "" Or TEXT3 = "" Or Text4 = "" Or Combo1 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.AddNew 'untuk tambah record'
Adodc1.Recordset.Fields("Foto") = Text7.Text
Adodc1.Recordset.Fields("NomorAnggota") = Text1.Text
Adodc1.Recordset.Fields("NamaLengkap") = Text2.Text
Adodc1.Recordset.Fields("Kelas") = TEXT3.Text
Adodc1.Recordset.Fields("JenisKelamin") = Combo1.Text
Adodc1.Recordset.Fields("NIS") = Text4.Text
Adodc1.Recordset.Fields("TempatTanggalLahir") = Text5.Text
Adodc1.Recordset.Fields("AlamatSekarang") = Text6.Text
Adodc1.Recordset!TanggalMulai = DTPicker1
Adodc1.Recordset.Update
'jika data berhaasil disimpan, tampilkan notif sukses
MsgBox ("Data berhasil disimpan!")
'panggil variabel untuk membersihkan form
Call bersih
Call query
Call KodeOtomatis
End If
End Sub

'Pindahkan data dari datagrid ke dalam form saat akan melakukan operasi, ubah, hapus atau cetak
Private Sub DataGrid1_Click()
Command2.Enabled = False
Text7.Text = Text8
Text1.Text = Adodc1.Recordset!NomorAnggota
Text2.Text = Adodc1.Recordset!NamaLengkap
TEXT3.Text = Adodc1.Recordset!Kelas
Combo1.Text = Adodc1.Recordset!JenisKelamin
Text4.Text = Adodc1.Recordset!NIS
Text5.Text = Adodc1.Recordset!TempatTanggalLahir
Text6.Text = Adodc1.Recordset!AlamatSekarang
DTPicker1 = Adodc1.Recordset!TanggalMulai
End Sub

'UBAH
'jika tombol ubah diklik
'script untuk merubah data anggota
Private Sub Command3_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text1 = "" Or Text2 = "" Or TEXT3 = "" Or Text4 = "" Or Combo1 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIUBAH !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi ubah
Adodc1.Recordset.Fields("Foto") = Text7.Text
Adodc1.Recordset.Fields("NomorAnggota") = Text1.Text
Adodc1.Recordset.Fields("NamaLengkap") = Text2.Text
Adodc1.Recordset.Fields("Kelas") = TEXT3.Text
Adodc1.Recordset.Fields("JenisKelamin") = Combo1.Text
Adodc1.Recordset.Fields("NIS") = Text4.Text
Adodc1.Recordset.Fields("TempatTanggalLahir") = Text5.Text
Adodc1.Recordset.Fields("AlamatSekarang") = Text6.Text
Adodc1.Recordset.Update
'jika data berhaasil diubah, tampilkan notif sukses
MsgBox ("Data berhasil diubah!")
Adodc1.Recordset!TanggalMulai = DTPicker1
'panggil variabel untuk membersihkan form
Call bersih
Call query
End If
End Sub

'HAPUS
'jika tombol hapus diklik
'script untuk hapus data anggota
Private Sub Command4_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text1 = "" Or Text2 = "" Or TEXT3 = "" Or Text4 = "" Or Combo1 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIHAPUS !", vbInformation, "PERHATIAN !"
Else
Dim pesan  As Integer
'tampilkan notifikasi pertanyaan
        pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini ?", vbCritical + vbYesNo, "WARNING !")
        If pesan = vbYes Then
        'jika user mengeklik "iya" maka hapus data
        Adodc1.Recordset.Delete
        'panggil variabel untuk membersihkan form
        Call bersih
        Call query
Else
End If
End If
End Sub

'CETAK DATA PERORANGAN
'jika tombol cetak perorangan diklik
Private Sub Command5_Click()
'etak berdasarkan data dalam form
cetak_anggota.Text1 = Text1.Text
cetak_anggota.Text2 = Text2.Text
cetak_anggota.TEXT3 = TEXT3.Text
cetak_anggota.Text4 = Text6.Text
cetak_anggota.Image1 = Image1.Picture
'Unload Me
cetak_anggota.Show
End Sub

'TAMPILKAN LAPORAN KESELURUHAN
Private Sub Command9_Click()
xx = "\LAP5.rpt"
cc = "*"
With crLAP5
    .ReportFileName = App.Path & xx
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End Sub

'BUKA FOTO
'jika tombol pilih foto diklik
Private Sub Command7_Click()
CommonDialog1.ShowOpen
'munculkan dialog pilih foto
Text7 = CommonDialog1.FileName
End Sub
Private Sub Text7_Change()
Image1.Picture = LoadPicture(Text7)
End Sub

'CARI
'jika tombol cari diklik
Private Sub Commandcari_Click()
If Text26.Text = "" Then
'tampilkan notif jika form pencaraian kososng
MsgBox "ISIKAN DATA PENCARIAN ANDA!", vbOKOnly, "Informasi!"
Else
'cari data berdasarkan nama atau nis
Adodc1.Recordset.Filter = "NamaLengkap like '%" + Me.Text26.Text + "%' or NomorAnggota like '%" + Me.Text26.Text + "%'"
End If
End Sub
'jika form pencaraian kosong refresh data
Private Sub Text26_Change()
If Text26.Text = "" Then
Call query
Else
'wkwk
End If
End Sub

'REFRESH TABEL
'jika tombol refresh ditekan
Private Sub Command8_Click()
Call query
Text26.Text = ""
End Sub


