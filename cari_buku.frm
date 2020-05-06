VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form cari_buku 
   BackColor       =   &H8000000B&
   Caption         =   "CARI BUKU"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text26 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   9480
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Commandcari 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "cari_buku.frx":0000
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   6376
      _Version        =   393216
      BackColor       =   16777088
      HeadLines       =   2
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7920
      Top             =   360
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
End
Attribute VB_Name = "cari_buku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FORM CARI DATA ANGGOTA
'MENAMPILKAN DATA ANGGOTA UNTUK MEMBANTU PENGISIAN FORM DALAM TRANSAKSI
'by INDRI DWI S
'======================================================================

'jika tombol cari di tekan
Private Sub Commandcari_Click()
'jika text 26 kosong tampilkan pesan notifikasi (tidak boleh kosong)
If Text26.Text = "" Then
MsgBox "ISIKAN DATA PENCARIAN ANDA!", vbOKOnly, "Informasi!"
Else
'jika text 26 tidak kosong
'mulai filter data berdasarkan nama atau nis
Adodc1.Recordset.Filter = "No_Induk_Buku like '%" + Me.Text26.Text + "%' or judul like '%" + Me.Text26.Text + "%'"
End If
End Sub

Private Sub DataGrid1_Click()
'pindaahkan data dari datagrid 1 data anggota, kedalam form transaksi pinjam untuk digunakan mengisi form
trans_pinjam.Text5.Text = Adodc1.Recordset!judul
'jika selesai tutup form ini
Unload Me
End Sub

Private Sub Form_Load()
'seting waktu di datagrid
With DataGrid1
.Columns(1).NumberFormat = "dd MMMM yy"
End With
End Sub

Private Sub Text26_Change()
'refresh data saat user sudah selesai memilih data
If Text26.Text = "" Then
Adodc1.Refresh
Else
'nothing
End If
End Sub

