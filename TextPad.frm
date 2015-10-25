VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_Main 
   Caption         =   "Quick ASM"
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   Picture         =   "TextPad.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOutputs 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   7275
   End
   Begin VB.TextBox Editor 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3705
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "TextPad.frx":0342
      Top             =   0
      Width           =   7230
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -15
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   2.54052e-29
   End
   Begin VB.TextBox Text_Error 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3840
      Width           =   7275
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu FileNew 
         Caption         =   "New"
      End
      Begin VB.Menu FileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu FileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu FileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "Edit"
      Begin VB.Menu EditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu EditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu EditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu EditSelect 
         Caption         =   "Select All"
      End
      Begin VB.Menu EditSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu EditFind 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu mnu_DEBUG 
      Caption         =   "&DEBUG"
      Enabled         =   0   'False
   End
   Begin VB.Menu ProcessMenu 
      Caption         =   "Process"
      Begin VB.Menu ProcessUpper 
         Caption         =   "Upper Case"
      End
      Begin VB.Menu ProcessLower 
         Caption         =   "Lower Case"
      End
   End
   Begin VB.Menu CustomMenu 
      Caption         =   "Customize"
      Begin VB.Menu CustomFont 
         Caption         =   "Font"
      End
      Begin VB.Menu CustomPage 
         Caption         =   "Page Color"
      End
      Begin VB.Menu CustomText 
         Caption         =   "Text Color"
      End
      Begin VB.Menu chmasmdir 
         Caption         =   "Change MASM directory"
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OpenFile As String                  ' holds the path of currently opened file
Dim mycaption As String                 ' holds current caption of the main window
Private WithEvents objDOS As DOSOutputs ' create object of class DOSOutputs. we are using this for creating hidden dos instance and executing masm commands.
Attribute objDOS.VB_VarHelpID = -1

Dim st As Integer
Dim filechanged As Boolean              ' flag to indicate if file in editor has been changed.

Private Sub about_Click()
frmAbout.Show (1)
End Sub

' to change the path of masm directory.

Private Sub chmasmdir_Click()
masmdir = GetSetting(App.EXEName, "config", "masmpath", "")         ' get previously saved settings from registry

masmdir = GetShortPath(InputBox("Enter Path of masm directory." & vbNewLine & "e.g. c:\masm", , masmdir))   ' get dos path from path entered by user
    
If masmdir = "" Then Exit Sub                           ' if canceled then exit from sub

SaveSetting App.EXEName, "config", "masmpath", masmdir  ' save path in registry.

End Sub

' to change font of editor
Private Sub CustomFont_Click()
On Error Resume Next
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    Editor.Font = CommonDialog1.FontName
    Editor.FontBold = CommonDialog1.FontBold
    Editor.FontItalic = CommonDialog1.FontItalic
    Editor.FontSize = CommonDialog1.FontSize
End Sub

' to change colour of editor
Private Sub CustomPage_Click()
    CommonDialog1.ShowColor
    Editor.BackColor = CommonDialog1.Color
End Sub

'to change font colour of editor
Private Sub CustomText_Click()
On Error Resume Next
    CommonDialog1.ShowColor
    Editor.ForeColor = CommonDialog1.Color
End Sub

' to copy selected text from editor.
Private Sub EditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Editor.SelText
End Sub

' to cut selected text from editor.
Private Sub EditCut_Click()
    Clipboard.SetText Editor.SelText
    Editor.SelText = ""
End Sub

' show find window
Private Sub EditFind_Click()
    Form_Search.Show
End Sub

Private Sub Editor_Change()
If Editor.Text = "" Then Exit Sub

' update caption
If OpenFile = "" Then
    mycaption = "Quick ASM Edit - Untitled#"
Else
    mycaption = "Quick ASM Edit - " & Split(OpenFile, "\")(UBound(Split(OpenFile, "\"))) & "#"
End If

filechanged = True      ' set changed flag.
End Sub

Private Sub Editor_Click()
If Editor.Text = "" Then Exit Sub

' update caption
If OpenFile = "" Then
    mycaption = "Quick ASM Edit - Untitled#"
Else
    mycaption = "Quick ASM Edit - " & Split(OpenFile, "\")(UBound(Split(OpenFile, "\"))) & "#"
End If

Me.Caption = mycaption & " - " & UBound(Split(Mid$(Editor.Text, 1, Editor.SelStart), vbNewLine)) + 1    ' disply current line number in form caption.
'***************************
livecompile                 ' call livecompile for error test in assembly
'***************************
End Sub


Private Sub Editor_KeyUp(KeyCode As Integer, Shift As Integer)
If Editor.Text = "" Then Exit Sub

If KeyCode = 116 And mnu_DEBUG.Enabled = True Then  ' if F5 shortcut key is pressed then open debug window.
    mnu_DEBUG_Click
End If

If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 13 Then    ' if up/down arrow key or enter pressed then update line number in window caption.
    Me.Caption = mycaption & " - " & UBound(Split(Mid$(Editor.Text, 1, Editor.SelStart), vbNewLine)) + 1
'***************************
    livecompile                 ' call livecompile for error test in assembly
'***************************
End If

On Error Resume Next
' adjust tab of new line to previous line. ( implement for space also .)
If KeyCode = 13 Then
    Dim tstr As String
    tstr = CStr(Split(Mid$(Editor.Text, 1, Editor.SelStart), vbNewLine)(UBound(Split(Mid(Editor.Text, 1, Editor.SelStart), vbNewLine)) - 1))
    Dim i As Integer, j As Integer
    For i = 1 To Len(tstr)
        If Mid(tstr, i, 1) = vbTab Then
            j = j + 1
        Else
            Exit For
        End If
    Next
    Editor.SelText = String(j, vbTab)
End If

' if error present then assemble everytime when key pressed.
If Not txtOutputs = "" Then
    If Not (KeyCode = 38 Or KeyCode = 40) Then
'***************************
        livecompile                 ' call livecompile for error test in assembly
'***************************
    End If
End If

End Sub
' to paste the copied text from keyboard.
Private Sub EditPaste_Click()
    If Clipboard.GetFormat(vbCFText) Then
        Editor.SelText = Clipboard.GetText
    Else
        MsgBox "Invalid Clipboard format."
    End If
End Sub
' to select all editor text.
Private Sub EditSelect_Click()
    Editor.SelStart = 0
    Editor.SelLength = Len(Editor.Text)
End Sub
' to exit program.
Private Sub FileExit_Click()
    End
End Sub
' to open new window ( not child in mdi )
Private Sub FileNew_Click()
If filechanged Then
    If MsgBox("Do you want to save changes made in current file ?", vbYesNo Or vbQuestion) = vbYes Then
        FileSave_Click  ' save changes
    End If
End If
On Error GoTo s1
Shell App.Path & "\" & App.EXEName, vbNormalFocus   ' open new instance of program
Exit Sub
' if new instance cannot be opened then renew current instance.
s1:
    Editor.Text = ""
    OpenFile = ""
End Sub
' to open file in editor
Private Sub FileOpen_Click()
' if previous file changed then ask user to save change.
If filechanged Then
    If MsgBox("Do you want to save changes made in current file ?", vbYesNo Or vbQuestion) = vbYes Then
        FileSave_Click
    End If
End If

' to open text file.
Dim FNum As Integer
Dim txt As String
On Error GoTo FileError
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.DefaultExt = "ASM"
    CommonDialog1.Filter = "ASM files|*.ASM|All files|*.*"
    CommonDialog1.ShowOpen
    FNum = FreeFile
    Open CommonDialog1.FileName For Input As #1
    txt = Input(LOF(FNum), #FNum)
    Close #FNum
    
    OpenFile = GetShortPath(CommonDialog1.FileName)
    Editor.Text = txt
    
    filechanged = False
    mycaption = "Quick ASM Edit - " & Split(OpenFile, "\")(UBound(Split(OpenFile, "\")))
    Me.Caption = mycaption
    
'***************************
    livecompile                 ' call livecompile for error test in assembly
'***************************
   
    Exit Sub
' show error if any ( try to fix - file with hidden artributes cannot be opened. )
FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while opening file " & CommonDialog1.FileName
    OpenFile = ""
    Close #FNum
End Sub
' to save text in editor.
Private Sub FileSave_Click()
Dim FNum As Integer
Dim txt As String

    If OpenFile = "" Then
        FileSaveAs_Click
        Exit Sub
    End If
On Error GoTo FileError
    FNum = FreeFile
    Open OpenFile For Output As #1
    Print #FNum, Editor.Text
    Close #FNum
    
    filechanged = False     ' reset file changed flag
    Exit Sub

FileError:
' show error if any.
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file " & OpenFile
    OpenFile = ""

End Sub
' to save file with diffrent name
Private Sub FileSaveAs_Click()
Dim FNum As Integer
Dim txt As String

On Error GoTo FileError
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    CommonDialog1.DefaultExt = "ASM"
    CommonDialog1.Filter = "ASM files|*.ASM|All files|*.*"
    CommonDialog1.ShowSave
    FNum = FreeFile
    Open CommonDialog1.FileName For Output As #1
    Print #FNum, Editor.Text
    Close #FNum
    OpenFile = GetShortPath(CommonDialog1.FileName)
    filechanged = False
    Exit Sub

FileError:
' show error if any.
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file " & CommonDialog1.FileName
    OpenFile = ""
End Sub

Private Sub Form_Load()
On Error GoTo errhandlter
Set objDOS = New DOSOutputs         ' intialise new DOSOutputs object

masmdir = GetSetting(App.EXEName, "config", "masmpath", "")     ' get previously saved masm directory path from registery if any.

If masmdir = "" Then    ' if no path is saved previously
    If CBool(PathFileExists(App.Path & "\masm.exe")) = False Then   ' check if masm exist in program directory
        masmdir = GetShortPath(InputBox("Enter Path of masm directory." & vbNewLine & "e.g. c:\masm"))  ' if not get dos path from user
    Else
        masmdir = GetShortPath(App.Path)    ' otherwise save program directory path
    End If
    
    If masmdir = "" Then
        MsgBox "I couldn't find MASM" & vbNewLine & "Exiting."
        End                                 ' if no masm directory found exit proram.
    End If
    SaveSetting App.EXEName, "config", "masmpath", masmdir  ' save masm directory path to registery for future use
End If

Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor    ' show version number on window capton.
Exit Sub
' show error if any.
errhandlter:
    MsgBox "Load Error " & vbNewLine _
        & Err.Description
    
End Sub
' adjust editor & error text sizes according to window size.
Private Sub Form_Resize()
On Error Resume Next
    Editor.Width = ScaleWidth
    txtOutputs.Width = ScaleWidth
    txtOutputs.Top = ScaleHeight - txtOutputs.Height
    Editor.Height = Form_Main.ScaleHeight - txtOutputs.Height
    
    Text_Error.Width = txtOutputs.Width
    Text_Error.Top = txtOutputs.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
If filechanged Then     ' if file has changed and not saved ask for saving before exit.
    If MsgBox("Do you want to save changes made in current file ?", vbYesNo Or vbQuestion) = vbYes Then
        FileSave_Click
    End If
End If
End Sub
' to open assebled & linked program in debug
Private Sub mnu_DEBUG_Click()
txtOutputs = ""
On Error GoTo errore
   
'***************************
livecompile                 ' call livecompile for error test in assembly
'***************************

If mnu_DEBUG.Enabled = False Then Exit Sub  ' if error exist in assebly exit sub

    FileSave_Click                          ' save opened file
    If OpenFile = "" Then Exit Sub          ' if file is not saved then exit sub
    
    objDOS.CommandLine = masmdir & "\masm.exe " & OpenFile      ' assemble current file using masm in hidden dos instance.
    objDOS.ExecuteCommand
        
    objDOS.CommandLine = masmdir & "\link.exe " & Split(Split(OpenFile, ".")(0), "\")(UBound(Split(Split(OpenFile, ".")(0), "\"))) & ".obj"     ' link assembled file created by masm ( with .obj extension ) using linker in hidden dos instance
    objDOS.ExecuteCommand
    
    Shell "debug " & masmdir & "\" & Split(Split(OpenFile, ".")(0), "\")(UBound(Split(Split(OpenFile, ".")(0), "\"))) & ".exe", vbNormalFocus   ' open linked file (.exe) in debug ( visible instance )
    
    Exit Sub
' show error if any
errore:
    MsgBox (" Error in  subDebug." & vbNewLine & Err.Description & " - " & Err.Source & " - " & CStr(Err.Number))

End Sub
' to change case of selected text in editor to lower.
Private Sub ProcessLower_Click()
Dim Sel1 As Integer, Sel2 As Integer
    
    Sel1 = Editor.SelStart
    Sel2 = Editor.SelLength
    Editor.SelText = LCase$(Editor.SelText)
    Editor.SelStart = Sel1
    Editor.SelLength = Sel2
End Sub
' to change case of selected text in editor to UPPER.
Private Sub ProcessUpper_Click()
Dim Sel1, Sel2 As Integer

    Sel1 = Editor.SelStart
    Sel2 = Editor.SelLength
    Editor.SelText = UCase$(Editor.SelText)
    Editor.SelStart = Sel1
    Editor.SelLength = Sel2
End Sub

Private Sub Text_Error_Change()
    Text_Error.SelStart = Len(Text_Error.Text)  ' seek at end of error output.
End Sub
' get dosoutput here
Private Sub objDOS_ReceiveOutputs(CommandOutputs As String)
    txtOutputs.Text = txtOutputs.Text & CommandOutputs
End Sub
' here we do following things :-
' save editor text in temp file
' assemble this file using masm in hidden dos instance
' Show error in error box if any
Public Sub livecompile()
' to save text in editor text box to temp file
Dim FNum As Integer
Dim txt As String

On Error GoTo FileError
    FNum = FreeFile
    Open masmdir & "\temp.asm" For Output As #1
    Print #FNum, Editor.Text
    Close #FNum

GoTo n1

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown error while saving file " & masmdir & "\temp.asm"
    OpenFile = ""
    
    
n1:
    txtOutputs.Text = ""            ' erase previous output from text box
    On Error GoTo errore
    Dim tcolor As Long

' give command to hidden dos to assemble text file we just saved.

    objDOS.CommandLine = masmdir & "\masm.exe " & masmdir & "\temp.asm"
    objDOS.ExecuteCommand
' after execution of commnad txtOutputs.Text contains output from hidden dos window

     Dim endpos As Integer, startpos As Integer
     endpos = InStr(1, txtOutputs.Text, "error")    ' check if dos output contains word "error"
     If endpos = 0 Then
         txtOutputs.Text = ""
         Text_Error.Text = ""
         mnu_DEBUG.Enabled = True
     Else
         mnu_DEBUG.Enabled = False                  ' disable debug menu if error exist
     End If
     
If Not Text_Error.Text = txtOutputs.Text Then            ' only update visible textbox if it is diiffrent from previous
    Text_Error.Text = txtOutputs.Text
End If

txtOutputs.SelStart = Len(txtOutputs.Text)          ' set cursor at end of text

    Exit Sub
    
' show error if any
errore:
    MsgBox (Err.Description & " - " & Err.Source & " - " & CStr(Err.Number))


End Sub
