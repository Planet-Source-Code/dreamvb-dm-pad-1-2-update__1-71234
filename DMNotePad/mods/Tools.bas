Attribute VB_Name = "Tools"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'Document helpers
Private Type TDocument
    OpenedDoc As String
    Changed As Boolean
End Type

Public Type TRGB
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Public mDocInfo As TDocument
Public dEditor As dTxtHelper

'Find Helpers
Public SelFindStr As String
Public SerOp As Integer

'File Filters
Public m_def_Filter As String
Public AppData As String
Public m_GotoOp As Boolean
Public m_CurSelPos As Long

Public ButtonPress As VbMsgBoxResult
Public Const VBQuote As String = """"

Public Function RgbToHex(r, g, b) As String
Dim WebColor As OLE_COLOR
    WebColor = b + 256 * (g + 256 * r)
    'Format Hex to 6 places
    RgbToHex = Right$("000000" & Hex$(WebColor), 6)
End Function

Public Sub Long2Rgb(lColor As Long, RgbType As TRGB)
Dim Tmp As Long
    Tmp = lColor
    'Convert Long To RGB
    With RgbType
        .Red = (Tmp Mod &H100)
        Tmp = (Tmp \ &H100)
        .Green = (Tmp Mod &H100)
        Tmp = (Tmp \ &H100)
        .Blue = (Tmp Mod &H100)
    End With
End Sub

Public Function GetFilename(lFile As String) As String
Dim sPos As Integer
    
    If Len(lFile) = 0 Then
        Exit Function
    Else
        sPos = InStrRev(lFile, "\", Len(lFile), vbBinaryCompare)
        If (sPos > 0) Then
            GetFilename = Mid(lFile, sPos + 1)
        Else
            GetFilename = lFile
        End If
    End If
End Function

Public Function GetFileTitle(lFile As String) As String
Dim sPos As Integer
    sPos = InStrRev(lFile, ".", Len(lFile), vbBinaryCompare)
    
    If (sPos > 0) Then
        GetFileTitle = Mid$(lFile, 1, sPos - 1)
    Else
        GetFileTitle = lFile
    End If
    
End Function

Public Function GetLeftSpace(SrcLine As String) As Integer
Dim x As Integer
Dim idx As Integer

    For x = 1 To Len(SrcLine)
        If Not IsWhite(Mid$(SrcLine, x, 1)) Then
            Exit For
        End If
    Next x
    
    GetLeftSpace = x
End Function

Public Function IsWhite(c As String) As Boolean
    If (c = vbTab) Or (c = " ") Then
        IsWhite = True
    Else
        IsWhite = False
    End If
End Function

Public Sub RunApp(iHwnd As Long, OpenOp As String, FileName As String)
Dim Ret As Long
    Ret = ShellExecute(iHwnd, OpenOp, FileName, "", "", 1)
End Sub

Public Sub CenterForm(frm As Form)
    frm.Top = (Screen.Height - frm.Height) \ 2
    frm.Left = (Screen.Width - frm.Width) \ 2
End Sub

Public Function FixPath(lPath As String) As String
    If Right(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Sub EnableBookmark(MnuItem As Menu, ByVal LineNumber As Long)
    If (Not MnuItem.Enabled) Then
        'Enable the menu if it not enabled.
        MnuItem.Enabled = True
    End If
    'Set bookmark Index
    MnuItem.Tag = LineNumber
End Sub

Public Function OpenFile(FileName As String) As String
Dim fp As Long
Dim Bytes() As Byte
    
    fp = FreeFile
    
    Open FileName For Binary As #fp
        If LOF(fp) > 0 Then
            ReDim Bytes(0 To LOF(fp) - 1)
            Get #fp, , Bytes
        End If
    Close #fp
    
    OpenFile = StrConv(Bytes, vbUnicode)
    Erase Bytes
End Function

Public Sub GetEnvironList(TMnu As Object)
Dim x As String
Dim Cnt As Integer
Dim sPos As Integer
Dim iCount As Integer
    
    'Add a menu of Environ variables.
    Cnt = 1
    x = Environ(Cnt)
    
    Do Until (x = vbNullString)
        x = Environ(Cnt)
        If Len(x) > 0 Then
            sPos = InStr(1, x, "=", vbBinaryCompare)
            
            If ((Cnt - 1) > 0) Then
                Load TMnu(Cnt - 1)
            End If
            'Set the menu caption
            TMnu(Cnt - 1).Visible = True
            TMnu(Cnt - 1).Caption = UCase$(Left$(x, sPos - 1))
        End If
        'INC Counter
        Cnt = (Cnt + 1)
    Loop
    
End Sub
