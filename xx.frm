VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57147393
      CurrentDate     =   43645
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57081857
      CurrentDate     =   43645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DTPicker2.Value = DTPicker1.Value + 7
End Sub

Private Sub Text1_Change()
On Error Resume Next


If Text1.Text = "" Then
DTPicker2.Value = DTPicker1.Value
End If
End Sub
