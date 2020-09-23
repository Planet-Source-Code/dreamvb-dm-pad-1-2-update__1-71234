VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtComment 
      Height          =   300
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   26
      Text            =   "#"
      Top             =   3570
      Width           =   720
   End
   Begin VB.TextBox txtChrWidth 
      Height          =   300
      Left            =   1200
      TabIndex        =   24
      Top             =   3225
      Width           =   720
   End
   Begin VB.TextBox txtTimeDate 
      Height          =   300
      Left            =   2355
      TabIndex        =   12
      Top             =   2925
      Width           =   2160
   End
   Begin VB.TextBox txtTabWidth 
      Height          =   300
      Left            =   1200
      TabIndex        =   11
      Top             =   2895
      Width           =   720
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5025
      Top             =   930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3975
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   3060
      TabIndex        =   13
      Top             =   4080
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4320
      TabIndex        =   15
      Top             =   4080
      Width           =   1110
   End
   Begin VB.CheckBox chkTabs 
      Caption         =   "Insert spaces as tabs"
      Height          =   225
      Left            =   2355
      TabIndex        =   9
      Top             =   2250
      Width           =   2865
   End
   Begin VB.CheckBox chkIdent 
      Caption         =   "Auto- indent"
      Height          =   225
      Left            =   2355
      TabIndex        =   8
      Top             =   1995
      Width           =   2865
   End
   Begin VB.TextBox txtMarin 
      Height          =   300
      Left            =   1200
      TabIndex        =   10
      Top             =   2565
      Width           =   720
   End
   Begin VB.CheckBox chkDrag 
      Caption         =   "Allow dragging and dropping of files"
      Height          =   225
      Left            =   2355
      TabIndex        =   7
      Top             =   1755
      Width           =   2865
   End
   Begin VB.CheckBox chkTray 
      Caption         =   "Minimize program to tray area"
      Height          =   225
      Left            =   2355
      TabIndex        =   5
      Top             =   1275
      Width           =   2910
   End
   Begin VB.CheckBox chkQExit 
      Caption         =   "Use Esc key for quick exit"
      Height          =   225
      Left            =   2355
      TabIndex        =   6
      Top             =   1515
      Width           =   2820
   End
   Begin VB.CheckBox chkMax 
      Caption         =   "Start Editor Maxsized"
      Height          =   225
      Left            =   2355
      TabIndex        =   4
      Top             =   1035
      Width           =   1905
   End
   Begin VB.Frame fraColors 
      Caption         =   "Colors"
      Height          =   1545
      Left            =   120
      TabIndex        =   14
      Top             =   945
      Width           =   2130
      Begin VB.PictureBox pBack 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   255
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   555
         Width           =   195
      End
      Begin VB.PictureBox pFore 
         BackColor       =   &H00000000&
         Height          =   195
         Left            =   255
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   285
         Width           =   195
      End
      Begin VB.Label lblBackCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         Height          =   195
         Left            =   510
         TabIndex        =   18
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lblForeColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ForeColor"
         Height          =   195
         Left            =   510
         TabIndex        =   17
         Top             =   285
         Width           =   675
      End
   End
   Begin VB.Frame frmFont 
      Caption         =   "Font"
      Height          =   765
      Left            =   105
      TabIndex        =   16
      Top             =   105
      Width           =   5355
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1065
      End
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   3765
      End
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   195
      Left            =   210
      TabIndex        =   25
      Top             =   3585
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chars Width:"
      Height          =   195
      Left            =   210
      TabIndex        =   23
      Top             =   3270
      Width           =   915
   End
   Begin VB.Label lblTimeDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time and date format"
      Height          =   195
      Left            =   2355
      TabIndex        =   22
      Top             =   2685
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tab Width:"
      Height          =   195
      Left            =   210
      TabIndex        =   20
      Top             =   2955
      Width           =   795
   End
   Begin VB.Label lblMargin 
      AutoSize        =   -1  'True
      Caption         =   "Left Margin:"
      Height          =   195
      Left            =   210
      TabIndex        =   19
      Top             =   2595
      Width           =   840
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function FindInList(Cbo As ComboBox, StrFind As String) As Integer
Dim x As Integer
Dim idx As Integer
    For x = 0 To Cbo.ListCount
        If LCase(StrFind) = LCase(Cbo.List(x)) Then
            idx = x
            Exit For
        End If
    Next x
    
    FindInList = idx
End Function

Private Function GetColorFromDLG() As Long
On Error GoTo CanErr:
    'Show color dialog.
    CD1.CancelError = True
    CD1.ShowColor
    GetColorFromDLG = CD1.Color
    Exit Function
CanErr:
    If (Err.Number = cdlCancel) Then
        GetColorFromDLG = -1
        Err.Clear
    End If
End Function
Private Function FixTextBox(sCode As Integer) As Integer

    'Fixes the textboxes to only allow digits and del key
    Select Case sCode
        Case 8
            FixTextBox = 8
        Case 48 To 57
            FixTextBox = sCode
        Case Else
            FixTextBox = 0
    End Select
End Function

Private Sub cmdOK_Click()
    'Save Config info
    SaveSetting "DMPad", "cfg", "FontName", cboFont.Text
    SaveSetting "DMPad", "cfg", "FontSize", cboSize.Text
    SaveSetting "DMPad", "cfg", "ForeColor", pFore.BackColor
    SaveSetting "DMPad", "cfg", "BackColor", pBack.BackColor
    SaveSetting "DMPad", "cfg", "Maxsized", chkMax.Value
    SaveSetting "DMPad", "cfg", "MoveToTray", chkTray.Value
    SaveSetting "DMPad", "cfg", "QuickExit", chkQExit.Value
    SaveSetting "DMPad", "cfg", "DragDropFiles", chkDrag.Value
    SaveSetting "DMPad", "cfg", "AutoIndent", chkIdent.Value
    SaveSetting "DMPad", "cfg", "UseSpace", chkTabs.Value
    SaveSetting "DMPad", "cfg", "Margin", Val(txtMarin.Text)
    SaveSetting "DMPad", "cfg", "TabWidth", Val(txtTabWidth.Text)
    SaveSetting "DMPad", "cfg", "TimeDate", txtTimeDate.Text
    SaveSetting "DMPad", "cfg", "CharWidth", txtChrWidth.Text
    SaveSetting "DMPad", "cfg", "Comment", txtComment.Text
    'Unload this form
    ButtonPress = vbOK
    Unload frmOptions
End Sub

Private Sub Command1_Click()
    ButtonPress = vbCancel
    Unload frmOptions
End Sub

Private Sub Form_Load()
Dim x As Integer
    Set frmOptions.Icon = Nothing
    
    'Setup font sizes
    For x = 8 To 72 Step 2
        cboSize.AddItem x
    Next x
    'Setup Font names
    For x = 0 To (Screen.FontCount - 1)
        cboFont.AddItem Screen.Fonts(x)
    Next x
    
    'Load Editor Config settings
    cboSize.ListIndex = FindInList(cboSize, GetSetting("DMPad", "cfg", "FontSize", 10))
    cboFont.ListIndex = FindInList(cboFont, GetSetting("DMPad", "cfg", "FontName", "Courier New"))
    chkMax.Value = GetSetting("DMPad", "cfg", "Maxsized", 0)
    chkTray.Value = GetSetting("DMPad", "cfg", "MoveToTray", 0)
    chkQExit.Value = GetSetting("DMPad", "cfg", "QuickExit", 0)
    chkDrag.Value = GetSetting("DMPad", "cfg", "DragDropFiles", 1)
    chkIdent.Value = GetSetting("DMPad", "cfg", "AutoIndent", 1)
    chkTabs.Value = GetSetting("DMPad", "cfg", "UseSpace", 0)
    txtMarin.Text = GetSetting("DMPad", "cfg", "Margin", 10)
    txtTabWidth.Text = GetSetting("DMPad", "cfg", "TabWidth", 8)
    pFore.BackColor = GetSetting("DMPad", "cfg", "ForeColor", 0)
    pBack.BackColor = GetSetting("DMPad", "cfg", "BackColor", vbWhite)
    txtTimeDate.Text = GetSetting("DMPad", "cfg", "TimeDate", "hh:mm dd/mm/yy")
    txtChrWidth.Text = GetSetting("DMPad", "cfg", "CharWidth", 80)
    txtComment.Text = GetSetting("DMPad", "cfg", "Comment", "#")
End Sub

Private Sub Form_Resize()
    Line3D1.Width = frmOptions.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub

Private Sub lblBackCol_Click()
    Call pBack_Click
End Sub

Private Sub lblForeColor_Click()
    Call pFore_Click
End Sub

Private Sub pBack_Click()
Dim sColor As Long
    'Set Picturebox with color
    sColor = GetColorFromDLG
    If (sColor <> -1) Then
        pBack.BackColor = sColor
    End If
End Sub

Private Sub pFore_Click()
Dim sColor As Long
    'Set Picturebox with color
    sColor = GetColorFromDLG
    If (sColor <> -1) Then
        pFore.BackColor = sColor
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = FixTextBox(KeyAscii)
End Sub

Private Sub txtMarin_KeyPress(KeyAscii As Integer)
    KeyAscii = FixTextBox(KeyAscii)
End Sub

Private Sub txtTabWidth_KeyPress(KeyAscii As Integer)
    KeyAscii = FixTextBox(KeyAscii)
End Sub
