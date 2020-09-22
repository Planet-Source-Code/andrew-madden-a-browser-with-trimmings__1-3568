VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHTML 
   Caption         =   "Andrew's Web Browser : Contact Information"
   ClientHeight    =   4005
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Left            =   480
      TabIndex        =   7
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   -2147483644
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   1
      Appearance      =   0
      TextRTF         =   $"frmHTML.frx":0000
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Text            =   "41701131"
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Text            =   "Or, you can reach him over ICQ at :"
      Top             =   3000
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Text            =   "thefc@hotmail.com"
      Top             =   2640
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Text            =   "email him at :"
      Top             =   2400
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000007&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Text            =   "To contact Andrew Madden"
      Top             =   2160
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close "
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3735
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4154
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "18/09/99"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "14:47"
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu menu 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmHTML"
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


Private Sub exit_Click()

    Unload Me

End Sub



Private Sub web_Click()
    frmBrowser.Show , Me
End Sub

