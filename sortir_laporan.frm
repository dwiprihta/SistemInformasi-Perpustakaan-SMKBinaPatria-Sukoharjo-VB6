VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form sortir_laporan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SORTIR LAPORAN"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "SORTIR LAPORAN"
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox cbulan 
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Text            =   "BULAN"
      Top             =   360
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "TAMPILKAN KESELURUHAN"
      Height          =   855
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ComboBox ctahun 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "TAHUN"
      Top             =   1080
      Width           =   6135
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   5760
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "sortir_laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Koneksi
RS.Open "select*from PERPUSTAKAAN where month(TanggalMasuk)='" & Val(cbulan) & "' and year(TanggalMasuk)='" & Val(ctahun) & "'", conn
If RS.EOF Then
MsgBox "DATA TIDAK DITEMUKAN !", vbInformation, "PERHATIAN !"

cbulan.SetFocus
Else
CR1.SelectionFormula = "Month({PERPUSTAKAAN.TanggalMasuk}) = " & Val(cbulan) & " And Year({PERPUSTAKAAN.TanggalMasuk}) = " & Val(ctahun) & ""
CR1.ReportFileName = App.Path & "\LAP2.rpt"
CR1.WindowState = crptMaximized
CR1.RetrieveDataFiles
CR1.Action = 1
End If
End Sub
   

Private Sub Command2_Click()
xx = "\LAP2.rpt"
cc = "*"
With CR1
    .ReportFileName = App.Path & xx
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End Sub

Private Sub Form_Load()
ctahun.AddItem ("2015")
ctahun.AddItem ("2016")
ctahun.AddItem ("2017")
ctahun.AddItem ("2018")
ctahun.AddItem ("2019")
ctahun.AddItem ("2020")
ctahun.AddItem ("2021")
ctahun.AddItem ("2022")
ctahun.AddItem ("2023")
ctahun.AddItem ("2024")
ctahun.AddItem ("2025")
ctahun.AddItem ("2026")
ctahun.AddItem ("2027")
ctahun.AddItem ("2028")
ctahun.AddItem ("2029")
ctahun.AddItem ("2030")

cbulan.AddItem ("1")
cbulan.AddItem ("2")
cbulan.AddItem ("3")
cbulan.AddItem ("4")
cbulan.AddItem ("5")
cbulan.AddItem ("6")
cbulan.AddItem ("7")
cbulan.AddItem ("8")
cbulan.AddItem ("9")
cbulan.AddItem ("10")
cbulan.AddItem ("11")
cbulan.AddItem ("12")

End Sub

