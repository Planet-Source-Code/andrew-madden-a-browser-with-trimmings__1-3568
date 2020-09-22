VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Andrew's Web Browser : HTML Editor"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5775
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10186
      _Version        =   393217
      TextRTF         =   $"frmEditor.frx":0000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Open"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Preview"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6495
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6588
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "23/07/99"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "18:45"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "&Open"
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
      End
      Begin VB.Menu menu 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_Click()

    frmAbout.Show , Me

End Sub

Private Sub Command1_Click()
    Unload Me

End Sub

Private Sub Command2_Click()
Open App.Path & "\preview.html" For Output As #1
Print #1, RichTextBox1.Text
Close #1
Load frmBrowser
frmBrowser.Show
frmBrowser.brwWebBrowser.Navigate App.Path & "\preview.html"

End Sub
Private Sub Command3_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, RichTextBox1.Text
    Close #1
End If

End Sub

Private Sub Command4_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do Until EOF(1)
    Line Input #1, lineoftext$
    alltext$ = alltext$ & lineoftext$
    RichTextBox1.Text = alltext$
    Loop
    Close #1
End If

End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
RichTextBox1.Text = "<HTML>" & vbCrLf & vbCrLf & "<HEAD>" & vbCrLf & "<TITLE>" & "Web Page</TITLE>" & vbCrLf & "</HEAD>" & vbCrLf & vbCrLf & "<BODY>" & vbCrLf & vbCrLf & "</BODY>" & vbCrLf & vbCrLf & "</HTML>" & vbCrLf & "<!--This was created using Andrew's Web Browser, and I'm proud of that fact>"

End Sub


Private Sub open_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As #1
    Do Until EOF(1)
    Line Input #1, lineoftext$
    alltext$ = alltext$ & lineoftext$
    RichTextBox1.Text = alltext$
    Loop
    Close #1
End If

End Sub


Private Sub save_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, RichTextBox1.Text
    Close #1
End If

End Sub
