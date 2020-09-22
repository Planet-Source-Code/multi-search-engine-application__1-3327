VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Teco Multi-Search Beta"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Lycos"
      Height          =   3855
      Left            =   6000
      TabIndex        =   8
      Top             =   4440
      Width           =   5775
      Begin SHDocVwCtl.WebBrowser WebBrowser4 
         Height          =   3495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5535
         ExtentX         =   9763
         ExtentY         =   6165
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   -1  'True
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Alta Vista"
      Height          =   3855
      Left            =   0
      TabIndex        =   7
      Top             =   4440
      Width           =   5895
      Begin SHDocVwCtl.WebBrowser WebBrowser3 
         Height          =   3495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5655
         ExtentX         =   9975
         ExtentY         =   6165
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   -1  'True
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Infoseek"
      Height          =   3615
      Left            =   6000
      TabIndex        =   6
      Top             =   720
      Width           =   5775
      Begin SHDocVwCtl.WebBrowser WebBrowser2 
         Height          =   3255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5535
         ExtentX         =   9763
         ExtentY         =   5741
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   -1  'True
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Yahoo!"
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   5895
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5655
         ExtentX         =   9975
         ExtentY         =   5530
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   -1  'True
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Navigation"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   255
         Left            =   7080
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtsearch 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Search For:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Try Adding Quotations Around Your Search Criteria To Get More Accurate Results"
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'               TECO'S MULTI-SEARCH
'****************************************************
' A simple application using 4 Webbrowsers, 1 Command
' button and 1 Text Box.  With this application you can
' see how easy it is to generate a Multi-Seach Engine
' Program.  Just replace the proper variables
' and viola.  This program is 100% Working.
'*****************************************************
' Created By:  TECO (aka T. Teed)
' Created In:  Visual Basic 5.0
' Created On:  August 29, 1999
' Contact Me:  teco@tecotown.com
'              http://www.tecotown.com
'******************************************************
'       ©1999 Tyler Teed.  All Rights Reserved.
'******************************************************

Private Sub Command1_Click()
'This is where we get the search criteria
'then we replace spaces with + signs
Let thesearch = txtsearch.Text
Dim i As intger
Let i = 1
While i <= Len(thesearch)
 If Mid(thesearch, i, 1) = " " Then
  Mid(thesearch, i, 1) = "+"
 End If
 i = i + 1
Wend
'Send Information to Webbrowsers and load criteria
WebBrowser1.Navigate "http://ink.yahoo.com/bin/query?p=" & thesearch & "&hc=0&hs=0"
WebBrowser2.Navigate "http://infoseek.go.com/Titles?qt=" & thesearch & "&col=WW&sv=IS&lk=noframes&svx=home_searchbox"
WebBrowser3.Navigate "http://www.altavista.com/cgi-bin/query?pg=q&kl=XX&stype=stext&q=" & thesearch
WebBrowser4.Navigate "http://www.lycos.com/cgi-bin/pursuit?cat=dir&query=" & thesearch
End Sub


Private Sub Form_Load()
'Initialize Each Browser's Search Engine
WebBrowser1.Navigate "http://www.yahoo.com"
WebBrowser2.Navigate "http://www.infoseek.com"
WebBrowser3.Navigate "http://www.altavista.com"
WebBrowser4.Navigate "http://www.lycos.com"
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Thank You and Goodbye Message / Contact Information
 msg = "Thank You For Using Teco Multi-Search"
 msg = msg & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Email:      teco@tecotown.com"
 msg = msg & Chr(13) & Chr(10) & "Website:  www.tecotown.com"
 sstyle = vbInformation
 ttitle = "Teco Multi-Search"
 MsgBox msg, sstyle, ttitle
 End
End Sub

Private Sub txtsearch_GotFocus()
'Highlight The Search Criteria for Easy Adjustments
txtsearch.SelStart = 0
txtsearch.SelLength = Len(txtsearch.Text)
End Sub
