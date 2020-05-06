VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form trans_selesai 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA TRANSAKSI SELESAI"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18765
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "selesai.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   18765
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text26 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   12600
      TabIndex        =   18
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   12135
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DataField       =   "Nama"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6960
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DataField       =   "NAP"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Kembali"
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
         Left            =   2400
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Pinjam"
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
         Left            =   360
         TabIndex        =   16
         Top             =   360
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
         Height          =   255
         Left            =   6960
         TabIndex        =   15
         Top             =   360
         Width           =   1215
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
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label8"
         DataField       =   "TanggalPinjam"
         DataSource      =   "Adodc1"
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
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label9"
         DataField       =   "TanggalKembali"
         DataSource      =   "Adodc1"
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
         Left            =   2400
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1920
      TabIndex        =   8
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin Crystal.CrystalReport crLAP4 
      Left            =   -120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H8000000D&
      Caption         =   "CETAK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   2235
   End
   Begin VB.TextBox Text3 
      DataField       =   "Kelas"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   8040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      DataField       =   "Alamat"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   1320
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   8640
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      DataField       =   "JudulBuku"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text5"
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   18000
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
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
      RecordSource    =   "PENGEMBALIAN"
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
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "HAPUS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   2250
   End
   Begin VB.CommandButton Command3 
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
      Height          =   555
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   1650
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "selesai.frx":0342
      Height          =   4575
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   8070
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   8760
      Width           =   2535
   End
End
Attribute VB_Name = "trans_selesai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'HAPUS
'jika tombol hapus dklik lakukan operasi hapus data
Private Sub Command2_Click()
'tampilkan peringatan jika form masih kosong
If Text1.Text = "" Then
MsgBox "Data pengembalian sudah kosong!", vbCritical, "WARNING !"
Else
        Dim pesan  As Integer
        'tampilkan pertanyaan hapus
        pesan = MsgBox("Apakah Anda yakin ingin menghapus data ini ?", vbCritical + vbYesNo, "WARNING !")
        'jika user menklik 'YA' hapus data
        If pesan = vbYes Then
        Adodc1.Recordset.Delete
        Else
        End If
End If
End Sub

'TAMPILKAN PENCARIAN TRANSAKSI SELESAI
Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox "ISIKAN DATA PENCARIAN ANDA!", vbOKOnly, "Informasi!"
Else
'cari data brdasarkan nama atau nomor anggota
Adodc1.Recordset.Filter = "Nama like '%" + Me.Text26.Text + "%' or NAP like '%" + Me.Text26.Text + "%'"
End If
End Sub

'jika tombol cetak diklik tampilkan laporan transaksi selesai
Private Sub Command6_Click()
xx = "\LAP4.rpt"
cc = "*"
With crLAP4
    '.SelectionFormula = "{PERPUSTAKAAN.No_Induk_Buku}='" & cc & "'"
    .ReportFileName = App.Path & xx
    .WindowState = crptMaximized
    '.Formulas(0) = "namakepsek='" & Label19.Caption & "'"
    .RetrieveDataFiles
    .Action = 1
End With
End Sub


'variabel yang dijalankan otomatis saat membuka form
Private Sub Form_Load()
With DataGrid1
.Columns(0).Width = 1800
.Columns(1).Width = 3600
.Columns(2).Width = 1000
.Columns(3).Width = 3400
.Columns(4).Width = 4500
.Columns(5).Width = 1800
.Columns(6).Width = 1800

.Columns(0).Caption = "NO ANGGOTA"
.Columns(1).Caption = "NAMA ANGGOTA"
.Columns(2).Caption = "KELAS/STATUS"
.Columns(3).Caption = "ALAMAT"
.Columns(4).Caption = "JUDUL BUKU"
.Columns(5).Caption = "TANGGAL PINJAM"
.Columns(6).Caption = "TANGGAL KEMBALI"

.Columns(5).NumberFormat = "dd MMMM yy"
.Columns(6).NumberFormat = "dd MMMM yy"
End With
End Sub

'refresh data saat pencarian berakhir
Private Sub Text26_Change()
If Text26.Text = "" Then
Adodc1.Refresh
With DataGrid1
.Columns(0).Width = 1800
.Columns(1).Width = 3600
.Columns(2).Width = 1000
.Columns(3).Width = 3400
.Columns(4).Width = 4500
.Columns(5).Width = 1800
.Columns(6).Width = 1800

.Columns(0).Caption = "NO ANGGOTA"
.Columns(1).Caption = "NAMA ANGGOTA"
.Columns(2).Caption = "KELAS/STATUS"
.Columns(3).Caption = "ALAMAT"
.Columns(4).Caption = "JUDUL BUKU"
.Columns(5).Caption = "TANGGAL PINJAM"
.Columns(6).Caption = "TANGGAL KEMBALI"


End With
Else
'wkwk
End If
End Sub
