VERSION 5.00
Begin VB.Form Form_Search 
   Caption         =   "Search & Replace"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form2"
   ScaleHeight     =   1980
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Case sensitive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   1260
      Width           =   2040
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5460
      TabIndex        =   8
      Top             =   1335
      Width           =   1245
   End
   Begin VB.CommandButton ReplaceAllButton 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5430
      TabIndex        =   7
      Top             =   630
      Width           =   1275
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "Replace"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4065
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find Again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5415
      TabIndex        =   5
      Top             =   120
      Width           =   1305
   End
   Begin VB.CommandButton FindButton 
      Caption         =   "Find"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4050
      TabIndex        =   4
      Top             =   105
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   105
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1485
      TabIndex        =   3
      Top             =   510
      Width           =   2130
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Find What"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   135
      Width           =   1320
   End
End
Attribute VB_Name = "Form_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Position As Integer

Private Sub FindButton_Click()
Dim compare As Integer

Position = 0
If Check1.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If
Position = InStr(Position + 1, Form_Main.Editor.Text, Text1.Text, compare)
If Position > 0 Then
    ReplaceButton.Enabled = True
    ReplaceAllButton.Enabled = True
    Form_Main.Editor.SelStart = Position - 1
    Form_Main.Editor.SelLength = Len(Text1.Text)
    Form_Main.SetFocus
Else
    MsgBox "String not found"
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
End If
End Sub

Private Sub FindNextButton_Click()
Dim compare As Integer

If Check1.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If
Position = InStr(Position + 1, Form_Main.Editor.Text, Text1.Text, compare)
If Position > 0 Then
    Form_Main.Editor.SelStart = Position - 1
    Form_Main.Editor.SelLength = Len(Text1.Text)
    Form_Main.SetFocus
Else
    MsgBox "String not found"
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
End If

End Sub

Private Sub Command5_Click()
    Form_Search.Hide
End Sub

Private Sub Form_Load()
Dim ret As Long
    Me.Show
    ret = SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.CurrentX, Me.CurrentY, Me.Width, Me.Height, SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Private Sub ReplaceButton_Click()
Dim compare As Integer

    Form_Main.Editor.SelText = Text2.Text
    If Check1.Value = 1 Then
        compare = vbBinaryCompare
    Else
        compare = vbTextCompare
    End If
    Position = InStr(Position + 1, Form_Main.Editor.Text, Text1.Text, compare)
    If Position > 0 Then
        Form_Main.Editor.SelStart = Position - 1
        Form_Main.Editor.SelLength = Len(Text1.Text)
        Form_Main.SetFocus
    Else
        MsgBox "String not found"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    
End Sub

Private Sub ReplaceAllButton_Click()
Dim compare As Integer

    Form_Main.Editor.SelText = Text2.Text
    If Check1.Value = 1 Then
        compare = vbBinaryCompare
    Else
        compare = vbTextCompare
    End If
    Position = InStr(Position + 1, Form_Main.Editor.Text, Text1.Text, compare)
    While Position > 0
        Form_Main.Editor.SelStart = Position - 1
        Form_Main.Editor.SelLength = Len(Text1.Text)
        Form_Main.Editor.SelText = Text2.Text
        Position = Position + Len(Text2.Text)
        Position = InStr(Position + 1, Form_Main.Editor.Text, Text1.Text)
    Wend
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
        MsgBox "Done replacing"
End Sub

