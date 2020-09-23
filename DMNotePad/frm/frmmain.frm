VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7455
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   Begin Project1.Tray Tray1 
      Left            =   6330
      Top             =   2385
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin MSComctlLib.Toolbar tBar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Plain Text"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Find"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Reduce Indent"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Increase Indent"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Wordwrap"
            ImageIndex      =   11
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6735
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":15EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1940
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2336
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2688
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":29DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pBase1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3870
      Width           =   3750
      Begin VB.ListBox LstOut 
         Appearance      =   0  'Flat
         Height          =   960
         IntegralHeight  =   0   'False
         Left            =   15
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   660
      End
      Begin VB.PictureBox pBar 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   15
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   168
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   2520
         Begin VB.Label lblOutput 
            AutoSize        =   -1  'True
            Caption         =   "Output"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   6
            Top             =   15
            Width           =   480
         End
         Begin VB.Image ImgBar 
            Height          =   105
            Left            =   30
            Picture         =   "frmmain.frx":2D2C
            Stretch         =   -1  'True
            Top             =   75
            Width           =   1185
         End
         Begin VB.Image ImgClose 
            Height          =   180
            Left            =   1515
            Picture         =   "frmmain.frx":2DA6
            ToolTipText     =   "Close"
            Top             =   30
            Width           =   210
         End
      End
   End
   Begin VB.PictureBox pBase 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   375
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      Width           =   5685
      Begin VB.PictureBox pLine 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1725
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox pLines 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   0
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   720
         Begin VB.Line lnMargin 
            BorderColor     =   &H00808080&
            X1              =   47
            X2              =   47
            Y1              =   0
            Y2              =   20
         End
      End
      Begin RichTextLib.RichTextBox TxtEdit 
         Height          =   795
         Left            =   705
         TabIndex        =   2
         Top             =   0
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1402
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         RightMargin     =   65000
         OLEDropMode     =   1
         TextRTF         =   $"frmmain.frx":2FF8
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6195
      Top             =   1695
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6795
      Top             =   1785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5070
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9111
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Begin VB.Menu mnutxtFile 
            Caption         =   "Plain Text"
         End
         Begin VB.Menu mnublank16 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNew2 
            Caption         =   "#"
            Index           =   0
         End
      End
      Begin VB.Menu mnuNewWindow 
         Caption         =   "New &Window..."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSet 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "&Email"
      End
      Begin VB.Menu mnublank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "#"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnublank12 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCpy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCpyAppend 
         Caption         =   "C&opy Append"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnublank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuMatch 
         Caption         =   "&Match Braket"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuMatch2 
         Caption         =   "Match Word"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "&Go To..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnublank4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDupLine 
         Caption         =   "Duplicate &Line"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuJoin 
         Caption         =   "Join Lines"
      End
      Begin VB.Menu mnublank13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuBlank18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComment 
         Caption         =   "Comment"
         Begin VB.Menu mnuComment1 
            Caption         =   "&Comment Block"
         End
         Begin VB.Menu mnuUnComment 
            Caption         =   "&UnComment Block"
         End
      End
      Begin VB.Menu mnublank14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelSave 
         Caption         =   "Sa&ve Selection..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "&Status Bar"
      End
      Begin VB.Menu mnuLineNum 
         Caption         =   "Line Numbers"
         Checked         =   -1  'True
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu mnuOutput 
         Caption         =   "Output"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuInst 
      Caption         =   "&Insert"
      Begin VB.Menu mnuDate 
         Caption         =   "Time/&Date"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuInstFile 
         Caption         =   "Insert &File"
      End
      Begin VB.Menu mnuFilePath 
         Caption         =   "Insert Filename &Path"
      End
      Begin VB.Menu mnuCurFile 
         Caption         =   "Current Filename"
      End
      Begin VB.Menu mnublank11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHexCol 
         Caption         =   "Insert &Hex Color"
      End
      Begin VB.Menu mnuUrl 
         Caption         =   "Insert &URL"
      End
      Begin VB.Menu mnuvbs 
         Caption         =   "&VBScript"
         Begin VB.Menu mnuEx 
            Caption         =   "Examples"
            Begin VB.Menu mnuExamples 
               Caption         =   "#"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuEnvA 
         Caption         =   "Environ Variable"
         Begin VB.Menu mnuEnvB 
            Caption         =   "#"
            Index           =   0
         End
      End
      Begin VB.Menu mnuBlank19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelimiter 
         Caption         =   "&Delimiters"
         Begin VB.Menu mnuD 
            Caption         =   "#"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuBook 
      Caption         =   "&Bookmark"
      Begin VB.Menu mnuMark1 
         Caption         =   "Mark1"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuMark2 
         Caption         =   "Mark2"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuMark3 
         Caption         =   "Mark3"
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuMark4 
         Caption         =   "Mark4"
         Enabled         =   0   'False
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuBlank9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSet1 
         Caption         =   "Set Mark1"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuSet2 
         Caption         =   "Set Mark2"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuSet3 
         Caption         =   "Set Mark3"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuSet4 
         Caption         =   "Set Mark4"
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu mnuConv 
      Caption         =   "&Convert"
      Begin VB.Menu mnuUpper 
         Caption         =   "To Uppercase"
      End
      Begin VB.Menu mnuLower 
         Caption         =   "To Lowercase"
      End
      Begin VB.Menu mnuTitleCase 
         Caption         =   "To Titlecase"
      End
      Begin VB.Menu mnuInvert 
         Caption         =   "In&vert Case"
      End
      Begin VB.Menu mnublank7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRott13 
         Caption         =   "ROT-1&3"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuRun2 
         Caption         =   "&Run Program"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MnuIE 
         Caption         =   "Run in IE"
      End
      Begin VB.Menu mnuExp 
         Caption         =   "E&xplorer"
      End
      Begin VB.Menu mnuCmdShell 
         Caption         =   "Command &Shell..."
      End
      Begin VB.Menu mnuCharMap 
         Caption         =   "Char&map"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnublank17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind2 
         Caption         =   "#"
         Index           =   0
      End
      Begin VB.Menu mnuBlank8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTool 
         Caption         =   "#"
         Index           =   0
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuWrap 
         Caption         =   "&Word Wrap"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuBlank15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPref 
         Caption         =   "&Preferences..."
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuMacro 
      Caption         =   "M&acro"
      Begin VB.Menu mnuRecord 
         Caption         =   "&Record Macro"
      End
      Begin VB.Menu mnuBlank10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMacros 
         Caption         =   "#"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp1 
         Caption         =   "View &Help"
      End
      Begin VB.Menu mnublank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About DMPad"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Msg1 = "Do you want to save the changes to "
Private Const MAX_RECENT As Integer = 5
Private mOldWinState As Integer
Private mFirstLoad As Boolean
Private mAutoIdent As Boolean
Private sMacroRecord As Boolean
Private sMacroFile As Long

Private Sub CommentBlock()
Dim Count As Integer
Dim vLines() As String
Dim sLine As String

    vLines = Split(TxtEdit.SelText, vbCrLf)
    For Count = 0 To UBound(vLines)
        sLine = GetSetting("DMPad", "cfg", "Comment", 0) & vLines(Count)
        vLines(Count) = sLine
    Next Count
    
    TxtEdit.SelText = Join(vLines, vbCrLf)
    'Clear up
    Erase vLines
    sLine = vbNullString
End Sub

Private Sub UnCommentBlock()
Dim Count As Integer
Dim vLines() As String
Dim sLine As String
Dim Tmp As String

    vLines = Split(TxtEdit.SelText, vbCrLf)
    For Count = 0 To UBound(vLines)
        sLine = vLines(Count)
        If Left(sLine, 1) = GetSetting("DMPad", "cfg", "Comment", 0) Then
            sLine = Right$(sLine, Len(sLine) - 1)
        End If
        vLines(Count) = sLine
    Next Count
    
    TxtEdit.SelText = Join(vLines, vbCrLf)
    'Clear up
    Erase vLines
    sLine = vbNullString
End Sub

Private Sub LoadDelimiterMenu()
Dim Count As Integer
Dim lFile As String
Dim sLine As String
Dim sPos As Integer
Dim fp As Long

    'Delimiters Filename
    lFile = AppData & "Delimiters.ini"
    'Hide first menu
    fp = FreeFile
    
    If Not FindFile(lFile) Then
        mnuDelimiter.Visible = False
        mnuBlank19.Visible = False
    End If
    
    Open lFile For Input As #fp
        Do Until EOF(fp)
            Line Input #fp, sLine
            sLine = Trim$(sLine)
            sPos = InStr(1, sLine, "=", vbBinaryCompare)
            
            If (sPos > 0) Then
                Count = (Count + 1)
                Load mnuD(Count)
                If LCase(sLine) = "spacer=" Then sLine = "-"
                
                mnuD(Count).Caption = Left$(sLine, sPos - 1)
                mnuD(Count).Visible = True
                mnuD(Count).Tag = Mid$(sLine, sPos + 1)
            End If
            
        Loop
    Close #fp
        
    mnuD(0).Visible = False
    
End Sub

Private Sub LoadTemplateMenu()
Dim Count As Integer
Dim xFile As String
Dim lPath As String
    'Template path
    lPath = AppData & "Templates\"

    'Hide first menu
    mnuNew2(0).Visible = False
    'Get First file
    xFile = Dir(lPath & "*.*")
    
    Do Until (xFile = "")
        'INC Counter
        Count = (Count + 1)
        'Load new menu item
        Load mnuNew2(Count)
        'Set menu info
        mnuNew2(Count).Visible = True
        mnuNew2(Count).Caption = GetFileTitle(xFile)
        mnuNew2(Count).Tag = lPath & xFile
        '
        tBar1.Buttons(1).ButtonMenus.Add , lPath & xFile, GetFileTitle(xFile)
        'Get next file
        xFile = Dir()
    Loop
    
    'Clear up
    lPath = vbNullString
    xFile = vbNullString
    
End Sub

Private Sub LoadFindMenu()
Dim fp As Long
Dim lFile As String
Dim sLine As String
Dim sPos As Integer
Dim Count As Integer

    fp = FreeFile
    lFile = AppData & "InetSerach.ini"
    'Hide first menu item
    mnuFind2(0).Visible = False
    'Check if the file is found
    If Not FindFile(lFile) Then
        mnublank17.Visible = False
        Exit Sub
    Else
        Count = 0
        Open lFile For Input As #fp
            Do Until EOF(fp)
                Line Input #fp, sLine
                sLine = Trim$(sLine)
                
                sPos = InStr(1, sLine, "=", vbBinaryCompare)
                If (sPos > 0) Then
                    Count = (Count + 1)
                    Load mnuFind2(Count)
                    'Update menu data
                    mnuFind2(Count).Caption = Left$(sLine, sPos - 1)
                    mnuFind2(Count).Tag = Trim$(Mid(sLine, sPos + 1))
                    mnuFind2(Count).Visible = True
                End If
            Loop
        Close #fp
    End If
    
    sLine = vbNullString
    
End Sub

Private Sub LoadVBSExamples()
Dim Count As Integer
Dim sLine As String
Dim fp As Long
Dim lFile As String
Dim sPos As Integer
    
    fp = FreeFile
    lFile = AppData & "VBScriptEx.ini"
    
    If Not FindFile(lFile) Then
        mnuvbs.Visible = False
        Exit Sub
    Else
        Count = 1
        'Open the examples file
        Open lFile For Input As #fp
            'Lopp tho the file
            Do Until EOF(fp)
                Line Input #fp, sLine
                sLine = Trim$(sLine)
                
                sPos = InStr(1, sLine, "=", vbBinaryCompare)
                If (sPos > 0) Then
                    Load mnuExamples(Count)
                    'Update menu caption
                    'This is used to add a Line in the menu
                    If LCase(Left$(sLine, sPos - 1)) = "spacer" Then
                        sLine = "-"
                    End If
                    
                    mnuExamples(Count - 1).Caption = Left$(sLine, sPos - 1)
                    'Update menu tag
                    mnuExamples(Count - 1).Tag = Mid$(sLine, sPos + 1)
                    'INC Count
                    Count = (Count + 1)
                End If
            Loop
        Close #fp
    End If
    
    'Hide last menu item
    mnuExamples(mnuExamples.Count - 1).Visible = False
    'Clear up
    sLine = vbNullString
    lFile = vbNullString
    
End Sub

Private Sub LoadConfig()
Dim DocChange As Boolean
    'Store the old doc changed
    DocChange = mDocInfo.Changed
    TxtEdit.Font.Name = GetSetting("DMPad", "cfg", "FontName", "Courier New")
    TxtEdit.Font.Size = GetSetting("DMPad", "cfg", "FontSize", 8)
    dEditor.ForeColor = GetSetting("DMPad", "cfg", "ForeColor", 0)
    TxtEdit.BackColor = GetSetting("DMPad", "cfg", "BackColor", vbWhite)
    TxtEdit.SelIndent = GetSetting("DMPad", "cfg", "Margin", 10)
    TxtEdit.OLEDropMode = GetSetting("DMPad", "cfg", "DragDropFiles", 1)
    'Restore doc changed
    mDocInfo.Changed = DocChange
End Sub

Private Sub LoadRecentMenu()
Dim Count As Integer
Dim lFile As String
Dim HasItems As Boolean

    For Count = 1 To MAX_RECENT
        lFile = GetSetting("DMPad", "Recent", Count, "")
        'Check for Length
        If Len(lFile) > 0 Then
            HasItems = True
            Call AddToRecentMenu(lFile)
        End If
    Next Count
    '
    mnublank12.Visible = HasItems
End Sub

Private Sub AddToRecentMenu(ByVal FileName As String)
Dim iCount As Integer
    'Check if the item is already in the menu
    For iCount = 0 To (mnuRecent.Count - 1)
        If LCase(FileName) = LCase(mnuRecent(iCount).Tag) Then
            Exit Sub
        End If
    Next iCount
    
    iCount = mnuRecent.Count
    
    If (iCount > 0) Then
        'Only load 5 items
        If (iCount <= MAX_RECENT) Then
            Load mnuRecent(iCount)
        End If
    End If
    
    If (iCount <= MAX_RECENT) Then
        mnuRecent(iCount).Visible = True
        mnuRecent(iCount).Caption = iCount & " " & GetFilename(FileName)
        mnuRecent(iCount).Tag = FileName
        'Save file to read edit
        SaveSetting "DMPad", "Recent", iCount, FileName
    Else
        mnuRecent(1).Caption = "1 " & GetFilename(FileName)
        mnuRecent(1).Tag = FileName
        'Save file to regedit.
        SaveSetting "DMPad", "Recent", "1", FileName
    End If
    
    'Show the spacer if we have items
    If (Not mnublank12.Visible) Then
        mnublank12.Visible = True
    End If
End Sub

Private Function ExecuteMacro(ByVal FileName As String) As Integer
Dim fp As Long
Dim sLine As String
Dim sCmd As String
Dim sPos As Integer
Dim sSelText As String
On Error Resume Next

    'This does the basic DMPad macro Lanuage.
    'Note it still in beta so it maybe buggy
    
    'Gets the selected text.
    sSelText = TxtEdit.SelText
    
    fp = FreeFile
    Open FileName For Input As #fp
        'Check if the file is empty
        If LOF(fp) = 0 Then
            'Close file
            Close #fp
            Exit Function
        End If
        'Loop tho the macro script.
        Do Until EOF(fp)
            Line Input #fp, sLine
            'Read in one line at a time.
            sLine = LTrim(sLine)
            'Check we have a length
            If Len(sLine) <> 0 Then
                'Check for : in the string
                sPos = InStr(1, sLine, ":", vbBinaryCompare)
                If (sPos > 0) Then
                    'extract command
                    sCmd = UCase$(Left$(sLine, sPos - 1))
                    'Fix the sLine so we only have the value
                    sLine = Mid$(sLine, sPos + 1)
                    'Phase commands
                    Select Case sCmd
                        Case "CUTLINE"
                            TxtEdit.SelLength = dEditor.LineLength
                            TxtEdit.SelText = ""
                        Case "COPYLINE"
                            Clipboard.SetText dEditor.GetLineText(dEditor.LineIndex), vbCFText
                        Case "CHAR"
                            'Insert a char
                            TxtEdit.SelText = Chr(Val(sLine))
                        Case "NEWLINE"
                            'Insert a new Line
                            TxtEdit.SelText = vbCrLf
                        Case "SELECTALL"
                            'Select all Text
                            Call dEditor.SelectAll
                        Case "COPY"
                            'Copy
                            dEditor.Copy
                        Case "CUT"
                            'Cut text
                            dEditor.Cut
                        Case "DELETE"
                            dEditor.Clear
                        Case "GOTOLINE"
                            'Goto a Line
                            dEditor.GotoLine = Val(sLine)
                        Case "SPACE"
                            'Inserts spaces
                            TxtEdit.SelText = sLine
                        Case "BACKSPACE"
                            'Delete
                            TxtEdit.SelStart = (TxtEdit.SelStart - 1)
                            TxtEdit.SelLength = 1
                            TxtEdit.SelText = ""
                        Case "SELMOVE"
                            'Moves to a cell position
                            TxtEdit.SelStart = Val(sLine) - 1
                        Case "INSERT"
                            TxtEdit.SelText = sLine
                        Case "INSERTDATE"
                            TxtEdit.SelText = Format(Now, GetSetting("DMPad", "cfg", "TimeDate", "hh:mm dd/mm/yy"))
                        Case "INSERTTIME"
                            TxtEdit.SelText = Time
                        Case "INSERTFILENAME"
                            TxtEdit.SelText = mDocInfo.OpenedDoc
                        Case "FILESIZE"
                            TxtEdit.SelText = Len(TxtEdit.Text)
                        Case "SETFOCUS"
                            TxtEdit.SetFocus
                        Case "SELTEXT"
                            TxtEdit.SelText = sSelText
                    End Select
                End If
            End If
        Loop
    Close #fp
    
    'Clear up
    ExecuteMacro = 1
    'sSelText = vbNullString
    sLine = vbNullString
    sCmd = vbNullString
    sPos = 0

End Function

Private Sub OpenMacro(ByVal FileName As String)
    sMacroFile = FreeFile
    Open FileName For Append As #sMacroFile
End Sub

Private Sub MacroAppend(ByVal SrcLine As String)
    Print #sMacroFile, SrcLine
End Sub

Private Sub CloseMacro()
    If (sMacroFile <> 0) Then
        Close #sMacroFile
        sMacroFile = 0
    End If
End Sub

Private Sub LoadMacroMenu()
Dim Count As Integer
Dim lFile As String
Dim lPath As String

    Count = mnuMacros.Count
    
    'Unload created menus
    Do Until (Count = 1)
        Count = (Count - 1)
        Unload mnuMacros(Count)
    Loop
    
    lPath = Dir(AppData & "Macros\*.ini")
    Count = 0
    
    Do Until Len(lPath) = 0
        'Load Menu Item
        If (Count > 0) Then
            Load mnuMacros(Count)
        End If
        'Set Items
        mnuMacros(Count).Caption = GetFileTitle(lPath)
        mnuMacros(Count).Tag = AppData & "Macros\" & lPath
        'INC Count
        Count = (Count + 1)
        lPath = Dir()
        DoEvents
    Loop
    
    mnuMacros(0).Visible = Count
    
End Sub

Private Sub LoadToolsMenu()
Dim lzFile As String
Dim iCount As Integer
Dim MyIni As New dINIFile

    lzFile = Dir(AppData & "Tools\*.ini")
    
    Do Until Len(lzFile) = 0
        'Set the INI Tool to read
        MyIni.FileName = AppData & "Tools\" & lzFile
        If (iCount > 0) Then
            'Only load if count is > 0
            Load mnuTool(iCount)
        End If
        'Show the menu item
        mnuTool(iCount).Visible = True
        'Set menu item caption
        mnuTool(iCount).Caption = MyIni.ReadValue("DMTool", "Caption")
        'Set menu tag with tools filename.
        mnuTool(iCount).Tag = MyIni.FileName
        'Inc menu count
        iCount = (iCount + 1)
        'Get next file
        lzFile = Dir()
    Loop
    
    mnuTool(0).Visible = iCount
    Set MyIni = Nothing
    'Clear up
    lzFile = vbNullString
    iCount = 0
End Sub

Public Sub DrawLines()
Dim counter As Long
Dim sLine As String

    'This sub draws the line numbers
    With pLines
        'Clear DC
        .Cls
        Set .Font = TxtEdit.Font
        For counter = (dEditor.VisableLine + 1) To dEditor.LineCount
            'Set normal text color
            .ForeColor = vbBlack
            .CurrentX = (.Width - 10) - .TextWidth(Str$(counter))
            If (counter = dEditor.LineIndex) Then
                'Set line heighlight color
                .ForeColor = &H808080
            End If
            'print lines
            
            pLines.Print counter
        Next counter
    End With
    
End Sub

Private Sub OpenDoc()
Dim lFile As String
    'Opens a new Text file and updates the editor.
    lFile = GetDLGName
    If Len(lFile) > 0 Then
        Call UpdateDoc(lFile)
    End If
End Sub

Private Sub UpdateDoc(ByVal FileName As String)
    Call dEditor.LoadFromFile(FileName)
    'Update forms caption
    frmmain.Caption = GetFilename(FileName) & " - DMPad"
    'Reset statusbar
    StatusBar1.Panels(2).Text = "Ln: 1 Col: 1 Sel: 0"
    'update open file
    mDocInfo.OpenedDoc = FileName
    'Update chnaged to false
    mDocInfo.Changed = False
    'Add to recent docs list
    Call AddToRecentMenu(FileName)
    'Check for .LOG
    If UCase(Left(TxtEdit.Text, 4)) = ".LOG" Then
        TxtEdit.SelStart = Len(TxtEdit.Text)
        'Add a blank line
        TxtEdit.SelText = vbCrLf
        'Insert Time and Date Log
        Call mnuDate_Click
    End If
End Sub

Private Sub NewDoc(Optional FileName As String = vbNullString, Optional Change As Boolean = False)
    
    If (Change) Then
        'Load template
        TxtEdit.Text = OpenFile(FileName)
    Else
        'Creates empty text document.
        TxtEdit.Text = FileName
    End If
    
    'Update forms caption.
    If (Not Change) Then
        frmmain.Caption = "Untitled - DMPad"
    End If
    
    'Update statusbar.
    StatusBar1.Panels(2).Text = "Ln: 1 Col: 1 Sel: 0"
    mDocInfo.OpenedDoc = vbNullString
    mDocInfo.Changed = Change
End Sub

Private Sub NewDocument(Optional FileName As String = vbNullString, Optional Change As Boolean = False)
Dim ans As Integer
Dim lFile As String
Dim lTmp As String
    
    lTmp = GetFilename(mDocInfo.OpenedDoc)
    
    If Len(lTmp) = 0 Then lTmp = "Untitled"
    
    If (mDocInfo.Changed) Then
        ans = MsgBox(Msg1 & lTmp & "?", vbYesNoCancel Or vbQuestion, "DMPad")
        'Check if No was pressed
        If (ans = vbNo) Then
            Call NewDoc(FileName, Change)
        ElseIf (ans = vbCancel) Then
           'Cancel pressed
        Else
            If Len(mDocInfo.OpenedDoc) > 0 Then
                'Save the document
                Call dEditor.SaveToFile(mDocInfo.OpenedDoc)
                'Create new document
                Call NewDoc(FileName, Change)
            Else
                'Show save dialog box.
                lFile = GetDLGName(False, "Save As")
                If (lFile = vbNullString) Then
                    'Cancel button was pressed
                Else
                    'Save the document.
                    Call dEditor.SaveToFile(lFile)
                    'Create new blank doc
                    Call NewDoc(FileName, Change)
                    lFile = vbNullString
                End If
            End If
        End If
    Else
        'Create new blank doc
        Call NewDoc(FileName, Change)
    End If
End Sub

Private Sub UpdateStatusbar()
    StatusBar1.Panels(2).Text = "Ln: " & dEditor.LineIndex & " Col: " & dEditor.LineLength & " Sel: " & TxtEdit.SelLength
End Sub

Private Sub EnableMenu1()
Dim HasSel As Boolean
    'Get Sel text.
    HasSel = Len(TxtEdit.SelText) > 0
    'Enable/Disable Menu items
    mnuCut.Enabled = HasSel
    mnuCpy.Enabled = HasSel
    mnuDelete.Enabled = HasSel
    mnuPaste.Enabled = dEditor.CanPaste
    mnuUndo.Enabled = dEditor.CanUndo
    mnuRedo.Enabled = dEditor.CanReDo
    mnuFindNext.Enabled = Len(TxtEdit.Text) > 0
    mnuMatch.Enabled = mnuFindNext.Enabled
    mnuMatch2.Enabled = HasSel 'mnuFindNext.Enabled
    mnuReplace.Enabled = mnuFindNext.Enabled
    mnuJoin.Enabled = HasSel
    MnuDupLine.Enabled = HasSel
    mnuSelSave.Enabled = HasSel
    mnuCpyAppend.Enabled = HasSel
    mnuComment.Enabled = HasSel
    
    'Toolbar Buttons
    tBar1.Buttons(6).Enabled = mnuCut.Enabled
    tBar1.Buttons(7).Enabled = mnuCut.Enabled
    tBar1.Buttons(8).Enabled = mnuPaste.Enabled
    tBar1.Buttons(9).Enabled = mnuFindNext.Enabled
    tBar1.Buttons(11).Enabled = mnuUndo.Enabled
    tBar1.Buttons(12).Enabled = dEditor.CanReDo
    tBar1.Buttons(14).Enabled = HasSel
    tBar1.Buttons(15).Enabled = HasSel
End Sub

Private Function GetDLGName(Optional ShowOpen As Boolean = True, Optional Title As String = "Open")
On Error GoTo CanErr:
        
    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = m_def_Filter
        .FilterIndex = 2
        
        If (ShowOpen) Then
            .ShowOpen
        Else
            .ShowSave
        End If
        
        GetDLGName = .FileName
        .FileName = vbNullString
    End With
    
    Exit Function
CanErr:

    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub Form_Activate()
    If Not (mFirstLoad) Then
        TxtEdit.SetFocus
        mFirstLoad = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Used for quick exit
    If (KeyCode = vbKeyEscape) Then
        'Call the menu exit sub
        If GetSetting("DMPad", "cfg", "QuickExit", "1") Then
            Call mnuExit_Click
        End If
    End If
End Sub

Private Sub Form_Load()
Dim sFile As String
On Error Resume Next

    'Setup tray control
    Tray1.ToolTip = frmmain.Caption
    Set Tray1.Icon = frmmain.Icon
    
    AppData = FixPath(App.Path) & "Data\"
    'Center this form
    Call CenterForm(frmmain)
    
    Set dEditor = New dTxtHelper
    dEditor.SetEditor = TxtEdit
    dEditor.GotoLine = 1
    
    Call mnuStatus_Click
    Call mnuOutput_Click
    Call GetEnvironList(mnuEnvB)
    Call EnableMenu1
    Call UpdateStatusbar
    Call LoadToolsMenu
    Call LoadMacroMenu
    Call LoadRecentMenu
    Call mnuLineNum_Click
    Call mnuWrap_Click
    Call LoadConfig
    Call LoadVBSExamples
    Call LoadFindMenu
    Call LoadTemplateMenu
    Call LoadDelimiterMenu
    
    frmmain.Caption = "Untitled - DMPad"
    mDocInfo.Changed = False
    
    sFile = Replace(Command$, Chr(34), "")
    
    If Len(sFile) > 0 Then
        Call dEditor.LoadFromFile(sFile)
        'Update forms caption
        frmmain.Caption = GetFilename(sFile) & " - DMPad"
        'Reset statusbar
        StatusBar1.Panels(2).Text = "Ln: 1 Col: 1 Sel: 0"
        'update open file
        mDocInfo.OpenedDoc = sFile
        'Update chnaged to false
        mDocInfo.Changed = False
    End If
    
    'File Filters
    m_def_Filter = "All Files (*.*)|*.*|Text Documents (*.txt,*.diz,*.nfo)|*.txt;*.diz;*.nfo" _
    & "|HTML Documents (*.html,*.htm,*.shtml)|*.html;*.htm;*.shtml" _
    & "|XML Documents (*.xml)|*.xml" _
    & "|C# Source Files (*.cs)|*.cs" _
    & "|VB.NET Source Files (*.vb,*.bas)|*.vb;*.bas;" _
    & "|VB Script (*.vbs)|*.vbs" _
    & "|Java Script (*.js)|*.js" _
    & "|Perl Scripts (*.pl,*.pm)|*.pl;*.pm;" _
    & "|C/C++ Source Files (*.cpp,*.c,*.h)|*.cpp;*.c;*.h" _
    & "|INI Files (*.ini,*.inf)|*.ini;*.inf;" _
    & "|Batch Files (*.bat)|*.bat" _
    & "|Active Server Pages (*.asp)|*.asp" _
    & "|PHP Documents (*.php,*.php3,*.php4,*.phtml)|*.php;*.ph34;*php4;*.phtml" _
    & "|Style Sheets (*.css)|*.css|" _
    & "Java (*.java)|*.java|"
    
    'Check if program is to open maxsized
    If GetSetting("DMPad", "cfg", "Maxsized", 0) Then
        frmmain.WindowState = 2
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim ans As Integer
Dim lFile As String
Dim lTmp As String

    lTmp = GetFilename(mDocInfo.OpenedDoc)
    If Len(lTmp) = 0 Then lTmp = "Untitled"
    
    If (mDocInfo.Changed) Then
        ans = MsgBox(Msg1 & lTmp & "?", vbYesNoCancel Or vbQuestion, "DMPad")
        'Check if No was pressed
        If (ans = vbCancel) Then
            Cancel = 1
            Exit Sub
        End If
        If (ans = vbYes) Then
            If Len(mDocInfo.OpenedDoc) > 0 Then
                'Save document
                Call dEditor.SaveToFile(mDocInfo.OpenedDoc)
            Else
                'Show save dialog box.
                lFile = GetDLGName(False, "Save As")
                If (lFile = vbNullString) Then
                    Cancel = 1
                Else
                    'Save the document.
                    Call dEditor.SaveToFile(lFile)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    If (frmmain.WindowState = 1) Then
        'Check if minsizeing to the tray
        If GetSetting("DMPad", "cfg", "MoveToTray", 0) = 1 Then
            Tray1.Visible = True
            frmmain.Visible = False
            frmmain.WindowState = 0
        End If
    End If
    
    'Resize editor
    pBase.Width = (frmmain.ScaleWidth - pBase.Left)
    pBase1.Width = pBase.Width
    
    If (mnuStatus.Checked) Then
        pBase1.Top = (frmmain.ScaleHeight - StatusBar1.Height - pBase1.Height)
    Else
        pBase1.Top = (frmmain.ScaleHeight - pBase1.Height) - 1
    End If
    
    If (Not mnuStatus.Checked) Then
        If (pBase1.Visible) Then
            pBase.Height = (frmmain.ScaleHeight - pBase.Top - pBase1.Height) - 1
        Else
            pBase.Height = (frmmain.ScaleHeight - pBase.Top - 1)
        End If
    Else
        If (pBase1.Visible) Then
            pBase.Height = (frmmain.ScaleHeight - StatusBar1.Height - pBase.Top - pBase1.Height) - 1
        Else
            pBase.Height = (frmmain.ScaleHeight - StatusBar1.Height - pBase.Top) - 1
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Derstory all the forms
    Set frmGoto = Nothing
    Set frmOptions = Nothing
    Set frmmain = Nothing
End Sub

Private Sub ImgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        Call mnuOutput_Click
    End If
End Sub

Private Sub mnuAbout_Click()
    ShellAbout frmmain.hwnd, frmmain.Caption, "Version 1.2" & vbCrLf & "By DreamVB", frmmain.Icon
End Sub

Private Sub mnuCharMap_Click()
    'Execure charmap program.
    Call RunApp(frmmain.hwnd, "open", "Charmap")
End Sub

Private Sub mnuCmdShell_Click()
Dim Ret As Long
    Call RunApp(frmmain.hwnd, "open", "cmd.exe")
End Sub

Private Sub mnuComment1_Click()
    Call CommentBlock
End Sub

Private Sub mnuCpy_Click()
    'Check if recording macro
    If (sMacroRecord) Then
        MacroAppend "Copy:"
    End If
    
    'Copy text to clipboard
    Call dEditor.Copy
End Sub

Private Sub mnuCpyAppend_Click()
Dim StrA As String
    'Copy Append
    StrA = Clipboard.GetText(vbCFText) & TxtEdit.SelText
    Clipboard.SetText StrA, vbCFText
End Sub

Private Sub mnuCurFile_Click()
    TxtEdit.SelText = mDocInfo.OpenedDoc
End Sub

Private Sub mnuCut_Click()
    
    'Check if recording macro
    If (sMacroRecord) Then
        MacroAppend "Cut:"
    End If
    
    'Cut text to clipboard
    Call dEditor.Cut
    'Update menu
    Call EnableMenu1
End Sub

Private Sub mnuD_Click(Index As Integer)
Dim sPos As Integer
Dim sLine As String
Dim sText As String

    sLine = mnuD(Index).Tag
    sPos = InStr(1, sLine, "$Txt", vbTextCompare)
    sText = TxtEdit.SelText
    
    If (sPos > 0) Then
        TxtEdit.SelText = Left$(sLine, sPos - 1) & sText & Mid$(sLine, sPos + 4)
    Else
        TxtEdit.SelText = sLine
    End If
    
End Sub

Private Sub mnuDate_Click()
    If (sMacroRecord) Then
        Call MacroAppend("InsertDate:")
    End If
    'Inset date and time.
    TxtEdit.SelText = Format(Now, GetSetting("DMPad", "cfg", "TimeDate", "hh:mm dd/mm/yy"))
End Sub

Private Sub mnuDelete_Click()
    'Check if recording macro
    If (sMacroRecord) Then
        MacroAppend "Delete:"
    End If
    'Clear Text
    Call dEditor.Clear
    'Update menu
    Call EnableMenu1
End Sub
Private Sub MnuDupLine_Click()
Dim StrA As String
    StrA = dEditor.GetLineText(dEditor.LineIndex)
    'Remove ctlf
    If Right(StrA, 2) = vbCrLf Then StrA = Left(StrA, Len(StrA) - 2)
    'Update editor.
    TxtEdit.Text = TxtEdit.Text & vbCrLf & StrA
    StrA = vbNullString
End Sub

Private Sub mnuEmail_Click()
    'Start up Email
    Call RunApp(frmmain.hwnd, "open", "mailto:yourname@yourmail.com?subject=Subject&body=" & TxtEdit.Text)
End Sub

Private Sub mnuEnvB_Click(Index As Integer)
    TxtEdit.SelText = "%" & mnuEnvB(Index).Caption & "%"
End Sub

Private Sub mnuExamples_Click(Index As Integer)
    TxtEdit.SelText = Replace(mnuExamples(Index).Tag, "|", vbCrLf) & vbCrLf
End Sub

Private Sub mnuExit_Click()
    Unload frmmain
End Sub

Private Sub mnuExp_Click()
    Call RunApp(frmmain.hwnd, "explore", "")
End Sub

Private Sub mnuFilePath_Click()
    'Insert Filename.
    TxtEdit.SelText = GetDLGName
End Sub

Private Sub mnuFind2_Click(Index As Integer)
Dim Tmp As String
    Tmp = mnuFind2(Index).Tag
    
    If (TxtEdit.SelLength > 0) Then
        Tmp = Replace(Tmp, "$find", TxtEdit.SelText, , , vbTextCompare)
        'Open IE
        Call RunApp(frmmain.hwnd, "open", Tmp)
        'Clear up
        Tmp = vbNullString
    End If
    
End Sub

Private Sub mnuFindNext_Click()
    SerOp = 0 'Show only find dialog.
    frmFind.Show , frmmain
End Sub

Private Sub mnuGoto_Click()
    'Get the current sel position
    m_CurSelPos = TxtEdit.SelStart
    'Set Goto Index
    dEditor.GotoLine = dEditor.LineIndex
    
    'Display goto line dialog.
    frmGoto.Show vbModal, frmmain
    If (ButtonPress = vbOK) Then
        If (m_GotoOp) Then
            If (dEditor.GotoLine > dEditor.LineCount) Then
                MsgBox "Line number out of range", vbExclamation, "Goto"
            End If
        Else
            'Goto Sel position
            TxtEdit.SelStart = m_CurSelPos
        End If
    End If
End Sub

Private Sub mnuHexCol_Click()
Dim MyRgb As TRGB
On Error GoTo CanErr:
    'Insert Hex Color
    With CD1
        .CancelError = True
        .ShowColor
        Call Long2Rgb(.Color, MyRgb)
        TxtEdit.SelText = VBQuote & "#" & RgbToHex(MyRgb.Red, MyRgb.Green, MyRgb.Blue) & VBQuote
    End With
    
    Exit Sub
CanErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Sub

Private Sub MnuIE_Click()
    If Len(mDocInfo.OpenedDoc) > 0 Then
        Call RunApp(frmmain.hwnd, "open", mDocInfo.OpenedDoc)
    End If
End Sub

Private Sub mnuInstFile_Click()
Dim lFile As String
    'Append file data to the document
    lFile = GetDLGName(, "Insert File")
    If Len(lFile) > 0 Then
        TxtEdit.SelText = OpenFile(lFile)
    End If
End Sub

Private Sub mnuInvert_Click()
    Call dEditor.Convert(dInvertCase)
End Sub

Private Sub mnuJoin_Click()
    Call dEditor.JoinLines
End Sub

Private Sub mnuLineNum_Click()
    'Hide or show line numbers
    mnuLineNum.Checked = (Not mnuLineNum.Checked)
    Timer1.Enabled = mnuLineNum.Checked
    pLines.Visible = mnuLineNum.Checked
    Call pBase_Resize
End Sub

Private Sub mnuLower_Click()
    Call dEditor.Convert(dLowerCase)
End Sub

Private Sub mnuMacros_Click(Index As Integer)
Dim Ret As Integer
    'Execute the macro
    Ret = ExecuteMacro(mnuMacros(Index).Tag)
    If (Ret <> 1) Then
        MsgBox "There was an error running the macro.", vbInformation, frmmain.Caption
    End If
End Sub

Private Sub mnuMark1_Click()
    dEditor.GotoLine = Val(mnuMark1.Tag)
End Sub

Private Sub mnuMark2_Click()
    dEditor.GotoLine = Val(mnuMark2.Tag)
End Sub

Private Sub mnuMark3_Click()
    dEditor.GotoLine = Val(mnuMark3.Tag)
End Sub

Private Sub mnuMark4_Click()
    dEditor.GotoLine = Val(mnuMark4.Tag)
End Sub

Private Sub mnuMatch_Click()
Dim sChr As String
Dim sStart As Integer
Dim X As Integer
Dim sEndChar As String

    sStart = TxtEdit.SelStart
    
    If (sStart > 0) Then
        sChr = Mid(TxtEdit.Text, TxtEdit.SelStart, 1)
        '
        If (sChr = "(") Then sEndChar = ")"
        If (sChr = "{") Then sEndChar = "}"
        If (sChr = "<") Then sEndChar = ">"
        If (sChr = "]") Then sEndChar = "["
            
        If (sChr Like "[({<[]") Then
            X = InStr(sStart, TxtEdit.Text, sEndChar, vbBinaryCompare)
            If (X > 0) Then
                TxtEdit.SelStart = (X - 1)
                TxtEdit.SelLength = 1
                TxtEdit.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub mnuMatch2_Click()
Dim sPos As Integer
Dim sText As String
    
    sText = TxtEdit.SelText
    sPos = InStr(TxtEdit.SelStart + Len(TxtEdit.SelText), TxtEdit.Text, TxtEdit.SelText, vbTextCompare)
    
    If (sPos > 0) Then
        TxtEdit.SelStart = (sPos - 1)
        TxtEdit.SelLength = Len(sText)
        TxtEdit.SetFocus
    End If
    
End Sub

Private Sub mnuNewWindow_Click()
    Call RunApp(frmmain.hwnd, "open", FixPath(App.Path) & App.EXEName)
End Sub

Private Sub mnuOpen_Click()
Dim ans As Integer
Dim lFile As String
Dim lTmp As String

    lTmp = GetFilename(mDocInfo.OpenedDoc)
    If Len(lTmp) = 0 Then lTmp = "Untitled"
    
    If (mDocInfo.Changed) Then
        ans = MsgBox(Msg1 & lTmp & "?", vbYesNoCancel Or vbQuestion, "DMPad")
        If (ans = vbNo) Then
            'Show open dialog.
            Call OpenDoc
        ElseIf (ans = vbCancel) Then
            'Cancel was pressed.
        Else
            'Button yes was presssed.
            If Len(mDocInfo.OpenedDoc) > 0 Then
                'Save the document
                Call dEditor.SaveToFile(mDocInfo.OpenedDoc)
                'Open the new document
                Call OpenDoc
            Else
                'show save dialog.
                lFile = GetDLGName(False, "Save As")
                If Len(lFile) > 0 Then
                    'Save document
                    Call dEditor.SaveToFile(lFile)
                    'Open new document
                    Call OpenDoc
                End If
            End If
        End If
    Else
        'Load the document.
        Call OpenDoc
    End If
    
End Sub

Private Sub mnuOutput_Click()
On Error Resume Next
    mnuOutput.Checked = (Not mnuOutput.Checked)
    pBase1.Visible = mnuOutput.Checked
    Call Form_Resize
End Sub

Private Sub mnuPageSet_Click()
On Error GoTo PrnErr:
    ' Show printer dialog
    With CD1
        .CancelError = True
        .DialogTitle = "Page Setup"
        .ShowPrinter
    End With
    
    Exit Sub
PrnErr:
    If (Err.Number <> cdlCancel) Then
        MsgBox Err.Description, vbCritical, "Error#" & Err.Number
    End If
End Sub

Private Sub mnuPaste_Click()
    'Paste text from clipboard
    If (dEditor.CanPaste) Then
        Call dEditor.Paste
    End If
End Sub

Private Sub mnuPref_Click()
    frmOptions.Show vbModal, frmmain
    
    If (ButtonPress = vbOK) Then
        Call LoadConfig
        If (Not mDocInfo.Changed) Then
            frmmain.Caption = Right(frmmain.Caption, Len(frmmain.Caption) - 1)
            Call pBase_Resize
        End If
    End If
End Sub

Private Sub mnuPrint_Click()
On Error GoTo PrnErr:

        With CD1
            .CancelError = True
            .DialogTitle = "Print"
            .Flags = (cdlPDReturnDC Or cdlPDNoPageNums)
            
            If (TxtEdit.SelLength = 0) Then
                .Flags = (.Flags Or cdlPDAllPages)
            Else
                .Flags = (.Flags Or cdlPDSelection)
            End If
            'Show print dialog.
            .ShowPrinter
            'Print document.
            Call TxtEdit.SelPrint(.hDC)
        End With
    
    Exit Sub
PrnErr:
    If (Err.Number <> cdlCancel) Then
        MsgBox Err.Description, vbCritical, "Error#" & Err.Number
    End If
End Sub

Private Sub mnuRecent_Click(Index As Integer)
Dim lFile As String
Dim ans As Integer
Dim lTmp As String
    
    'Temp filename.
    lFile = mnuRecent(Index).Tag
    'Check if the filename is found
    
    If Not FindFile(lFile) Then
        MsgBox lFile & vbCrLf & vbCrLf & "Was not found.", vbInformation, "File Not Found"
        Exit Sub
    End If
    
    lTmp = GetFilename(mDocInfo.OpenedDoc)
    If Len(lTmp) = 0 Then lTmp = "Untitled"
    
    If (mDocInfo.Changed) Then
        ans = MsgBox(Msg1 & lTmp & "?", vbYesNoCancel Or vbQuestion, "DMPad")
        'Check if no was pressed
        If (ans = vbNo) Then
            'Open the file.
            Call UpdateDoc(lFile)
        End If
        'Check if Cancel was pressed
        If (ans = vbCancel) Then
            Exit Sub
        End If
        'Check if yes was pressed.
        If (ans = vbYes) Then
            'Check if we already have a file open.
            If Len(mDocInfo.OpenedDoc) > 0 Then
                Call dEditor.SaveToFile(mDocInfo.OpenedDoc)
                Call UpdateDoc(lFile)
            Else
                'show save dialog.
                lFile = GetDLGName(False, "Save As")
                If Len(lFile) > 0 Then
                    'Save document
                    Call dEditor.SaveToFile(lFile)
                    lFile = mnuRecent(Index).Tag
                    Call UpdateDoc(lFile)
                End If
            End If
        End If
    Else
        'Open the file.
        Call UpdateDoc(lFile)
    End If
    
End Sub

Private Sub mnuRecord_Click()
Dim sName As String
    
    sMacroRecord = (Not sMacroRecord)
    
    If (sMacroRecord) Then
        'Ask user to enter a name
        sName = Trim$(InputBox$("Please enter a menu item name:", "Macro 1"))
        'Check for name
        If Len(sName) = 0 Then
            sMacroRecord = False
        Else
            'Record a new macro
            Call OpenMacro(AppData & "Macros\" & sName & ".ini")
            mnuRecord.Caption = "&Stop Macro"
            TxtEdit.MousePointer = rtfCustom
        End If
    Else
        mnuRecord.Caption = "&Record Macro"
        TxtEdit.MousePointer = rtfDefault
        'Stop macro
        Call CloseMacro
        'Add the macro to the menu
        Call LoadMacroMenu
    End If
End Sub

Private Sub mnuRedo_Click()
    Call dEditor.Undo
    tBar1.Buttons(11).Enabled = True
    tBar1.Buttons(12).Enabled = False
    mnuRedo.Enabled = False
    mnuUndo.Enabled = True
End Sub

Private Sub mnuReplace_Click()
    SerOp = 1 'Shows find and replace dialog.
    frmFind.Show , frmmain
End Sub

Private Sub MnuRott13_Click()
    Call dEditor.Convert(dROT13)
End Sub

Private Sub mnuRun2_Click()
On Error Resume Next
Dim oShellApp As Object
    Set oShellApp = CreateObject("Shell.Application")
    Call oShellApp.filerun
End Sub

Private Sub mnuSave_Click()

    If (mDocInfo.OpenedDoc = vbNullString) Then
        Call mnuSaveAs_Click
    Else
        'Save the document
        Call dEditor.SaveToFile(mDocInfo.OpenedDoc)
        'Check if document changed
        If (mDocInfo.Changed) Then
           'Update Programs Caption
           frmmain.Caption = Right$(frmmain.Caption, Len(frmmain.Caption) - 1)
        End If
        'Set document changed to false.
        mDocInfo.Changed = False
    End If
End Sub

Private Sub mnuSaveAs_Click()
    mDocInfo.OpenedDoc = GetDLGName(False, "Save As")
    
    If (mDocInfo.OpenedDoc <> vbNullString) Then
        'Save the document
        Call dEditor.SaveToFile(mDocInfo.OpenedDoc)
        'Check if document changed
        If (mDocInfo.Changed) Then
           'Update Programs Caption
           frmmain.Caption = Right$(frmmain.Caption, Len(frmmain.Caption) - 1)
        End If
        'Set document changed to false.
        mDocInfo.Changed = False
    End If
End Sub

Private Sub mnuSelAll_Click()
    'Check if recording macro
    If (sMacroRecord) Then
        MacroAppend "SelectAll:"
    End If
    
    'Select all Text
    Call dEditor.SelectAll
    'Update menu
    Call EnableMenu1
End Sub

Private Sub mnuSelSave_Click()
Dim lFile As String
    lFile = GetDLGName(False, "Save Selection")
    If Len(lFile) > 0 Then
        'Save the selection
        Call dEditor.SaveSelection(lFile)
    End If
End Sub

Private Sub mnuSet1_Click()
    Call EnableBookmark(mnuMark1, dEditor.LineIndex)
End Sub

Private Sub mnuSet2_Click()
    Call EnableBookmark(mnuMark2, dEditor.LineIndex)
End Sub

Private Sub mnuSet3_Click()
    Call EnableBookmark(mnuMark3, dEditor.LineIndex)
End Sub

Private Sub mnuSet4_Click()
    Call EnableBookmark(mnuMark4, dEditor.LineIndex)
End Sub

Private Sub mnuStatus_Click()
    mnuStatus.Checked = (Not mnuStatus.Checked)
    StatusBar1.Visible = mnuStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuTitleCase_Click()
    Call dEditor.Convert(dTitleCase)
End Sub

Private Sub mnuTool_Click(Index As Integer)
Dim sCmd As String
Dim isDone As Boolean
Dim DosOut As New DosStdOut
Dim MyIni As New dINIFile
Dim vLines() As String
Dim Count As Integer
Dim sLine As String

    'Get the tools filename
    MyIni.FileName = mnuTool(Index).Tag
    sCmd = MyIni.ReadValue("DMTool", "Cmd")
    
    If Len(mDocInfo.OpenedDoc) = 0 Then
        Call mnuSave_Click
    End If
    
    'Check if user did save the file
    If Len(mDocInfo.OpenedDoc) = 0 Then
        Exit Sub
    Else
        If MyIni.ReadValue("DMTool", "SaveFile", "0") = 1 Then
            'Save the current file
            If (mDocInfo.Changed) Then
                Call mnuSave_Click
            End If
        End If
        'Replace the $File var with the current filename if present
        sCmd = Replace(sCmd, "$File", mDocInfo.OpenedDoc, , vbTextCompare)
        sCmd = Replace(sCmd, "$Path", FixPath(App.Path), , vbTextCompare)

        'Execute the command
        If MyIni.ReadValue("DMTool", "DosPipe", "0") = 1 Then
            If DosOut.DosPipe(sCmd, 256) Then
                'Check if we are sending to the console.
                If MyIni.ReadValue("DMTool", "Showoutput", "0") = 1 Then
                    'Clear listbox.
                    LstOut.Clear
                    'Show output window is not visable
                    If (Not pBase1.Visible) Then
                        pBase1.Visible = True
                    End If
                    'Store the lines in vlines array.
                    vLines = Split(DosOut.Outputs, vbCrLf)
                    For Count = 0 To UBound(vLines)
                        'Extract the line to add.
                        sLine = Trim(vLines(Count))
                        If Len(sLine) > 1 Then
                            'Display outputs in the listbox
                            LstOut.AddItem vLines(Count)
                        End If
                    Next Count
                End If
            End If
        Else
            'Execute normal command
            'Call DosOut.DosPipe(sCmd, 256)
            RunApp frmmain.hwnd, "open", sCmd
        End If
    End If
    
    'Clear up
    Set DosOut = Nothing
    Set MyIni = Nothing
    sCmd = vbNullString
    sLine = vbNullString
    Erase vLines
End Sub

Private Sub mnuToolbar_Click()
    mnuToolbar.Checked = (Not mnuToolbar.Checked)
    tBar1.Visible = mnuToolbar.Checked
    
    'Hide / show toolbar
    If (tBar1.Visible) Then
        pBase.Top = (tBar1.Height + 2)
    Else
        pBase.Top = 0
    End If
    
    Call Form_Resize
End Sub

Private Sub mnutxtFile_Click()
    Call NewDocument
End Sub

Private Sub mnuUnComment_Click()
    Call UnCommentBlock
End Sub

Private Sub mnuUndo_Click()
    'undo chnages
    Call dEditor.Undo
    '
    tBar1.Buttons(11).Enabled = False
    tBar1.Buttons(12).Enabled = True
    mnuRedo.Enabled = True
    mnuUndo.Enabled = False
End Sub

Private Sub mnuUpper_Click()
    Call dEditor.Convert(dUpperCase)
End Sub

Private Sub mnuUrl_Click()
Dim StrA As String
Dim StrB As String

    StrA = InputBox("Enter the URL Address", "Insert URL", "http://www.somesite.com")
    StrB = InputBox("Enter a name for the URL", "Insert Name")
    'Insert the url
    TxtEdit.Text = "<a href=" & VBQuote & StrA & VBQuote & ">" & StrB & "</a>"
End Sub

Private Sub mnuWrap_Click()
    'Turn on / Off wrapping
    mnuWrap.Checked = (Not mnuWrap.Checked)
    tBar1.Buttons(17).Value = Abs(mnuWrap.Checked)
    Call dEditor.WordWrap(mnuWrap.Checked)
End Sub

Private Sub pBar_Resize()
On Error Resume Next
    ImgClose.Left = (pBar.ScaleWidth - ImgClose.Width) - 6
    ImgBar.Width = (pBar.ScaleWidth - ImgClose.Width) - 12
End Sub

Private Sub pBase_Resize()
On Error Resume Next
Dim X1 As Long
Dim X2 As Long
Dim cW As Long

    If (mnuLineNum.Checked) Then
        TxtEdit.Left = pLines.ScaleWidth - 1
    Else
        TxtEdit.Left = 0
    End If
    
    pLines.Height = pBase.Height
    lnMargin.Y2 = pLines.ScaleHeight
    TxtEdit.Width = (pBase.ScaleWidth - TxtEdit.Left)
    TxtEdit.Height = pBase.ScaleHeight
    
    With pLine
        cW = GetSetting("DMPad", "cfg", "CharWidth", 80)
        .Height = TxtEdit.Height - 16
        X1 = pLines.TextWidth("W")
        X2 = pLines.TextWidth("i")
        If (X1 = X2) Then
        .Left = (cW * X2 - 1)
        End If
    End With
    
End Sub

Private Sub pBase1_Resize()
    LstOut.Width = pBase.ScaleWidth + 3
    pBar.Width = pBase.Width
End Sub

Private Sub pLines_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Goto the line number the user clicked on.
    dEditor.GotoLine = (Y \ pLines.TextHeight("Zx") + 1)
End Sub

Private Sub tBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            'New
            'Call mnuNew_Click
        Case 2
            'Open
            Call mnuOpen_Click
        Case 3
            'Save
            Call mnuSave_Click
        Case 4
            'Print
            Call mnuPrint_Click
        Case 6
            'Cut
            Call mnuCut_Click
        Case 7
            'Copy
            Call mnuCpy_Click
        Case 8
            'Paste
            Call mnuPaste_Click
        Case 9
            'Find Next
            Call mnuFindNext_Click
        Case 11
            mnuUndo_Click
        Case 12
            'Redo
            Call mnuRedo_Click
        Case 14
            'UnIndent
            Call dEditor.UnIndent
        Case 15
            'Indent
            Call dEditor.Indent
        Case 17
            'Word wrap
            Call mnuWrap_Click
    End Select
End Sub

Private Sub tBar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If (ButtonMenu.Index = 1) Then
        Call mnutxtFile_Click
    Else
        If FindFile(ButtonMenu.Key) Then
            Call NewDocument(ButtonMenu.Key, True)
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Call DrawLines
End Sub

Private Sub Tray1_MouseUp(Button As Integer)
On Error Resume Next
    If (Button = vbLeftButton) Then
        Tray1.Visible = False
        frmmain.Visible = True
    End If
End Sub

Private Sub TxtEdit_Change()
Dim sTmp As String

    If (Not mDocInfo.Changed) Then
        sTmp = "*" & frmmain.Caption
        'Document was chnaged.
        frmmain.Caption = sTmp
        mDocInfo.Changed = True
    End If

    'Call DrawLines
    Call EnableMenu1
    Call UpdateStatusbar
End Sub

Private Sub TxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sLine As String
Dim X As Integer
Dim sSpace As String
    
    'Check if paste key is used.
    If (KeyCode = vbKeyV) And (Shift = 2) Then
        Call mnuPaste_Click
        'Set keycode to zero
        KeyCode = 0
    End If
    
    'Check if recording macro
    If (sMacroRecord) Then
        'Up,Down,Left,Right keys used for macros
        If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Or (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then
            Call MacroAppend("SelMove:")
        End If
    End If
    
    'Check for tab key
    If (KeyCode = 9) Then
        'Check if using spaces instade of tabs
        If GetSetting("DMPad", "cfg", "UseSpace", 0) Then
            TxtEdit.SelText = Space$(GetSetting("DMPad", "cfg", "TabWidth", 8))
            KeyCode = 0
        End If
    End If
    
    'Check if auto ident is enabled.
    If GetSetting("DMPad", "cfg", "AutoIndent", 1) Then
        If (KeyCode = 13) Then
            'Get current line
            sLine = dEditor.GetLineText(dEditor.LineIndex)
            
            'Get the number of Tabs
            X = GetLeftSpace(sLine)
            'Indent spaceing
            sSpace = Left(sLine, X - 1)
            '
            If (X <> 1) Then
                'Pad out with the number of ident space's
                TxtEdit.SelText = vbCrLf & sSpace
                'Check if recording macro
                If (sMacroRecord) Then
                    'Add Line command
                    Call MacroAppend("NewLine:")
                    Call MacroAppend("Space:" & sSpace)
                End If
                KeyCode = 0
                sSpace = vbNullString
            End If
        End If
    End If
End Sub

Private Sub TxtEdit_KeyPress(KeyAscii As Integer)
On Error Resume Next
    'Check if recording macro
    If (sMacroRecord) Then
        If (KeyAscii = 8) Then
            'Add delete command.
            'sMacro.Add "Backspace:"
            Call MacroAppend("Backspace:")
            Exit Sub
        ElseIf (KeyAscii = 13) Then
            'Add new line command
            'sMacro.Add "NewLine:"
            Call MacroAppend("NewLine:")
        Else
            'Add chars
            Call MacroAppend("Char:" & KeyAscii)
        End If
    End If
End Sub

Private Sub TxtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UpdateStatusbar
    Call EnableMenu1
End Sub

Private Sub TxtEdit_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lFile As String
Dim ans As Integer
Dim lTmp As String
    
    'Temp filename.
    lFile = Data.Files(1)
    lTmp = GetFilename(mDocInfo.OpenedDoc)
    
    If Len(lTmp) = 0 Then lTmp = "Untitled"
    
    If (mDocInfo.Changed) Then
        ans = MsgBox(Msg1 & lTmp & "?", vbYesNoCancel Or vbQuestion, "DMPad")
        'Check if no was pressed
        If (ans = vbNo) Then
            'Open the file.
            Call UpdateDoc(lFile)
        End If
        'Check if Cancel was pressed
        If (ans = vbCancel) Then
            Exit Sub
        End If
        'Check if yes was pressed.
        If (ans = vbYes) Then
            'Check if we already have a file open.
            If Len(mDocInfo.OpenedDoc) > 0 Then
                Call dEditor.SaveToFile(mDocInfo.OpenedDoc)
                Call UpdateDoc(lFile)
            Else
                'show save dialog.
                lFile = GetDLGName(False, "Save As")
                If Len(lFile) > 0 Then
                    'Save document
                    Call dEditor.SaveToFile(lFile)
                    lFile = Data.Files(1)
                    Call UpdateDoc(Data.Files(1))
                End If
            End If
        End If
    Else
        'Open the file.
        Call UpdateDoc(lFile)
    End If
    
End Sub

Private Sub TxtEdit_SelChange()
    SelFindStr = TxtEdit.SelText
    Call EnableMenu1
    Call UpdateStatusbar
End Sub
