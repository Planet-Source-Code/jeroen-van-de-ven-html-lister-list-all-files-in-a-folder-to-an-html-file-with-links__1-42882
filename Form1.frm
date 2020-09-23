VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "HTML Lister 2003 v0.1"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Text            =   "*.*"
      Top             =   2640
      Width           =   975
   End
   Begin RichTextLib.RichTextBox HTML 
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create HTML"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2520
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "Form1.frx":0082
      Left            =   120
      List            =   "Form1.frx":0084
      TabIndex        =   4
      Top             =   3000
      Width           =   8175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Pattern:"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   2480
      Width           =   975
   End
   Begin VB.Line Line5 
      X1              =   4800
      X2              =   4800
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line Line4 
      X1              =   3600
      X2              =   3600
      Y1              =   2520
      Y2              =   2880
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   8280
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8280
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8280
      Y1              =   2460
      Y2              =   2460
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Y
Public S As String

Private Sub Command1_Click()
Dim X
List1.Clear
For X = 0 To File1.ListCount - 1
List1.AddItem File1.List(X)
Pauze
If List1.ListCount = File1.ListCount Then MsgBox "DONE!", , "Html Lister"
Next X
End Sub

Private Sub Command2_Click()
HTML.Text = "<!DOCTYPE=FILE_LIST!><!GENERATOR=HTML LISTER!><html><head><title>Download MP3's!</title></head>" & vbCrLf & vbCrLf & "<body bgcolor='#8795F1'>" & vbCrLf & vbCrLf & "<p>Bestanden/Files: " & List1.ListCount & "</p>" & vbCrLf & vbCrLf
For Y = 0 To List1.ListCount - 1
Repl
Pauze
HTML.Text = HTML.Text & "<p><a href='\" & S & "'>" & List1.List(Y) & "</a></p>" & vbCrLf
Pauze
Next Y
HTML.Text = HTML.Text & vbCrLf & "</body>" & vbCrLf & vbCrLf & "</html>"
Open File1.Path & "\index.html" For Output As #1
Print #1, HTML.Text
Close #1
MsgBox "Done!", , "Html Lister"
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
File1.Path = Drive1.Drive
Dir1.Path = Drive1.Drive
End Sub
Sub Repl()
S = Replace(Replace(Replace(Replace(Replace(Replace(Replace(List1.List(Y), " ", "%20"), "!", "%21"), "#", "%23"), "+", "%2b"), ",", "%2c"), "-", "%2d"), "'", "%27")
End Sub
Sub Pauze()
Dim Z
For Z = 0 To 10000
If Z = 10000 Then Exit Sub
Next Z
End Sub



Private Sub Form_Resize()
On Error GoTo 8
Line1.X2 = Me.Width - 240
Line2.X2 = Me.Width - 240
Line3.X2 = Me.Width - 240
List1.Width = Me.Width - 360
File1.Width = Me.Width - 3240
HTML.Width = Me.Width - 360
HTML.Height = Me.Height - 4500
8:
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then File1.Pattern = Text1.Text
End Sub
