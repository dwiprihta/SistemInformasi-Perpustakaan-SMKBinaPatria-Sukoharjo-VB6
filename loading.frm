VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form loading 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1755
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   10920
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "PERPUSTAKAAN BINA PATRIA"
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
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORM LOADING
'by INDRI DWI S
'======================================================================

Private Sub Timer1_Timer()
        With ProgressBar1
        If .Value < 100 Then
        DoEvents
        .Value = .Value + 1
        If .Value = 2 Then
        'Label1.Caption = "Loading . . .": DoEvents
        End If
If .Value = 100 Then
frmLogin.Show
Unload Me
End If
        End If
        End With
End Sub
