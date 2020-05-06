VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form data_buku 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA BUKU"
   ClientHeight    =   11055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20370
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "data_buku.frx":0000
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3360
      TabIndex        =   37
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   158728193
      CurrentDate     =   43534
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13920
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   3600
      Width           =   4455
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13920
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   3120
      Width           =   4455
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13920
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   2640
      Width           =   4455
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13920
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13920
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13920
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Left            =   13920
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3480
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   2880
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1080
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Refresh"
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
      Left            =   18600
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton CommandTambah 
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
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Commandoke 
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
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Commandcari 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cari"
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
      Left            =   17160
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   1335
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   240
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
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
      RecordSource    =   "PERPUSTAKAAN"
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1200
      TabIndex        =   14
      Top             =   4680
      Width           =   9255
      Begin VB.CommandButton Command8 
         BackColor       =   &H8000000D&
         Caption         =   "Cetak Laporan"
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
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
         Height          =   495
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Commandubah 
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
         Height          =   495
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H80000003&
      Height          =   525
      Left            =   13800
      TabIndex        =   13
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   19920
      Top             =   120
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "data_buku.frx":3354
      Height          =   5055
      Left            =   360
      TabIndex        =   36
      Top             =   5760
      Width           =   19695
      _ExtentX        =   34740
      _ExtentY        =   8916
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
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   19560
      TabIndex        =   18
      Top             =   10800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Total buku    :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   17760
      TabIndex        =   17
      Top             =   10800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Katalog"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12240
      TabIndex        =   12
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12240
      TabIndex        =   11
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12240
      TabIndex        =   10
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Buku"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12240
      TabIndex        =   9
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sumber Dana"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12240
      TabIndex        =   8
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Editor Cetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12240
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tahun Terbit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12240
      TabIndex        =   6
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Kota Terbit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Penerbit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pengarang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Judul"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Induk Buku"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Masuk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "data_buku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FORM DATA BUKU
'MENAMPILKAN DATA BUKU DAN OPERASI (SIMPAN, HAPUS, UBAH, DAN CETAK)
'by INDRI DWI S
'======================================================================

'FORM LOAD
'perintah otomatis yang dijalankan saat form data anggota dibuka

Sub Form_Load()
'seting waktu di datagrid
With DataGrid1
.Columns(1).NumberFormat = "dd MMMM yy"
End With
'panggil variabel untuk memebersihkan form
Call KodeOtomatis
Call bersih
End Sub

'kode anggota otomatis
Sub KodeOtomatis()
Call Koneksi
RS.Open ("select * from PERPUSTAKAAN Where No_Induk_Buku In(Select Max(No_Induk_Buku)From PERPUSTAKAAN)Order By No_Induk_Buku Desc"), conn
RS.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RS
        If .EOF Then
            Urutan = "BKU" + "001"
            Text1 = Urutan
        Else
            Hitung = Right(!No_Induk_Buku, 3) + 1
            Urutan = "BKU" + Right("000" & Hitung, 3)
        End If
        Text1 = Urutan
    End With
End Sub

'BERSIH
'membuat variabel untuk membersihkan data pada form (dipanggil pada tombol tambah, edit, hapus)
Sub bersih()
'Text1.Text = ""
Text2.Text = ""
TEXT3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
DTPicker1.Value = Now
End Sub

'ENABEL
'membuat variabel untuk membuat form menjadi hidup (dipanggil pada tombol tambah)
Sub enabel()
'Text1.Enabled = True
Text2.Enabled = True
TEXT3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text2.SetFocus
End Sub

'REFRESH TABEL
'membuat variabel untuk merefresh data (dipanggil pada tombol tambah, edit, hapus)
Private Sub Command2_Click()
Adodc1.Refresh
'=====================================
With DataGrid1
.Columns(0).Width = 1450
.Columns(1).Width = 1650
.Columns(2).Width = 3700
.Columns(3).Width = 1200
.Columns(4).Width = 1400
.Columns(5).Width = 1400
.Columns(6).Width = 1340
.Columns(7).Width = 1300
.Columns(8).Width = 1400
.Columns(9).Width = 1200
.Columns(10).Width = 900
.Columns(11).Width = 1500
.Columns(12).Width = 1120
End With
'=====================================
End Sub

'TAMBAH
'jika tombol tambah dikllik
Private Sub Commandtambah_Click()
Commandoke.Enabled = True
Call bersih
Call enabel
Call KodeOtomatis
End Sub

'PINDAH DATA DARI DGV KE TEXTBOX
'jika dtagrid diklik pindahkan data pada datagrid ke form
Private Sub DataGrid1_Click()
Commandoke.Enabled = False
Text1.Text = Adodc1.Recordset!No_Induk_Buku
DTPicker1 = Adodc1.Recordset!TanggalMasuk
Text2.Text = Adodc1.Recordset!judul
TEXT3.Text = Adodc1.Recordset!Pengarang
Text4.Text = Adodc1.Recordset!Penerbit
Text5.Text = Adodc1.Recordset!Kota_Terbit
Text6.Text = Adodc1.Recordset!Tahun_Terbit
Text7.Text = Adodc1.Recordset!Editor_Cetak
Text8.Text = Adodc1.Recordset!Sumber_Dana
Text9.Text = Adodc1.Recordset!JumlahBuku
Text10.Text = Adodc1.Recordset!Harga
Text11.Text = Adodc1.Recordset!Keterangan
Text12.Text = Adodc1.Recordset!Katalog
End Sub

'SIMPAN
'jika tombol simpan diklik
'script untuk menyimpan data buku
Private Sub CommandOke_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text1 = "" Or Text2 = "" Or TEXT3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN !", vbInformation, "PERHATIAN !"
Else
'jika tidak ada inputan yang kosog simpan data buku
Adodc1.Recordset.AddNew 'untuk tambah record'
Adodc1.Recordset!No_Induk_Buku = Text1.Text
Adodc1.Recordset!TanggalMasuk = DTPicker1
Adodc1.Recordset!judul = Text2.Text
Adodc1.Recordset!Pengarang = TEXT3.Text
Adodc1.Recordset!Penerbit = Text4.Text
Adodc1.Recordset!Kota_Terbit = Text5.Text
Adodc1.Recordset!Tahun_Terbit = Text6.Text
Adodc1.Recordset!Editor_Cetak = Text7.Text
Adodc1.Recordset!Sumber_Dana = Text8.Text
Adodc1.Recordset!JumlahBuku = Text9.Text
Adodc1.Recordset!Harga = Text10.Text
Adodc1.Recordset!Keterangan = Text11.Text
Adodc1.Recordset!Katalog = Text12.Text
Adodc1.Recordset.Update
'jika berhasil disimpan tampilkan pesan sukses
MsgBox "Data sudah disimpan!", vbOKOnly, "Informasi!"
'panggil variabel untyk membersihkan form
Call bersih
Call KodeOtomatis
End If
End Sub

'UBAH
'jika tombol ubah diklik
'script untuk merubah data buku
Private Sub Commandubah_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text1 = "" Or Text2 = "" Or TEXT3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA UBAH !", vbInformation, "PERHATIAN !"
Else
'jika tidak ada inputan yang kosog ubah data buku
Adodc1.Recordset!No_Induk_Buku = Text1.Text
Adodc1.Recordset!TanggalMasuk = DTPicker1
Adodc1.Recordset!judul = Text2.Text
Adodc1.Recordset!Pengarang = TEXT3.Text
Adodc1.Recordset!Penerbit = Text4.Text
Adodc1.Recordset!Kota_Terbit = Text5.Text
Adodc1.Recordset!Tahun_Terbit = Text6.Text
Adodc1.Recordset!Editor_Cetak = Text7.Text
Adodc1.Recordset!Sumber_Dana = Text8.Text
Adodc1.Recordset!JumlahBuku = Text9.Text
Adodc1.Recordset!Harga = Text10.Text
Adodc1.Recordset!Keterangan = Text11.Text
Adodc1.Recordset!Katalog = Text12.Text
Adodc1.Recordset.Update
'jika berhasil diubah tampilkan pesan sukses
MsgBox "Data sudah diubah!", vbOKOnly, "Informasi!"
'panggil variabel untyk membersihkan form
Call bersih
End If
End Sub

'HAPUS
'jika tombol hapus diklik
'script untuk hapus data buku
Private Sub Command1_Click()
If Text1 = "" Then
'jika tidak ada inputan yang kosog ubah data buku
MsgBox "PILIH DAHULU DATA YANG AKAN DIHAPUS ", vbInformation, "PERHATIAN !"
Else
Dim pesan  As Integer
        'tampilkan pesan pertanyaan
        pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini ?", vbCritical + vbYesNo, "WARNING !")
        If pesan = vbYes Then
        'jika usser mengkili 'iya' maka hapus data
        Adodc1.Recordset.Delete
        'panggil variabel untyk membersihkan form
        Call bersih
Else
End If
End If
End Sub

'CARI
'script untuk pencarian data
Private Sub Commandcari_Click()
If Text26.Text = "" Then
MsgBox "ISIKAN DATA PENCARIAN ANDA!", vbOKOnly, "Informasi!"
Else
'cari berdasarkan no buku atau juduk
Adodc1.Recordset.Filter = "Judul like '%" + Me.Text26.Text + "%' or No_Induk_Buku like '%" + Me.Text26.Text + "%'"
End If
End Sub
'refresh data pencarain
Private Sub Text26_Change()
If Text26.Text = "" Then
Adodc1.Refresh
Else
'nothing
End If
End Sub

'CETAK LAPORAN BUKU
'jika tomobol cetak diklik tampilkanlaporan data buku
Private Sub Command8_Click()
sortir_laporan.Show
End Sub


