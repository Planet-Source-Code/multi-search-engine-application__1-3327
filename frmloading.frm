VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmloading 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Teco Multi-Search"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2400
      Top             =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INITIALIZING SEARCH ENGINES"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmloading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
'Start the Timer to control the Progress Bar
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
'Increase the % complete of progress bar
ProgressBar1.Value = ProgressBar1.Value + 5
If ProgressBar1.Value = 45 Then
 'when we reach 45% then load Form1 into memory
 Load Form1
End If
If ProgressBar1.Value = 100 Then
 'when 100% complete, hide this form
 frmloading.Visible = False
 'show the main form
 Form1.Visible = True
 'disable this timer
 Timer1.Enabled = False
End If
End Sub
