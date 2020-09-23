VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPropertyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''
''''
''''Private Const LVM_FIRST As Long = &H1000
''''' Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)        '
''''' Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57) '
''''' Private Const LVHT_NOWHERE As Long = &H1                    '
''''' Private Const LVHT_ONITEMICON As Long = &H2                 '
''''' Private Const LVHT_ONITEMLABEL As Long = &H4                '
''''' Private Const LVHT_ONITEMSTATEICON As Long = &H8            '
''''
''''
''''Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
''''' Private Const LVS_EX_SUBITEMIMAGES As Long = &H2                      '
''''' Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54) '
''''' Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55) '
''''
''''
''''' Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or _
''''LVHT_ONITEMLABEL Or _
''''        LVHT_ONITEMSTATEICON)
''''
''''Private Const LVIR_BOUNDS = 0
''''Private Const LVIR_ICON = 1
''''Private Const LVIR_LABEL = 2
''''Private Const LVIR_SELECTBOUNDS = 3
''''
''''
''''
''''Private Const LVMODE_TEXT = 0
''''Private Const LVMODE_COMBO = 1
''''Private Const LVMODE_BUTTON = 2
''''
''''
''''' Private Type POINTAPI '
''''' X As Long             '
''''' Y As Long             '
''''' End Type              '
''''
''''Private Type Rect
''''    Left As Long
''''    Top As Long
''''    Right As Long
''''    Bottom As Long
''''End Type
''''
''''' Private Type LVHITTESTINFO                         '
''''' pt As POINTAPI                                     '
''''' flags As Long                                      '
''''' iItem As Long                                      '
''''' iSubItem  As Long  'ie3+ only .. was NOT in win95. '
''''' 'Valid only for LVM_SUBITEMHITTEST                 '
''''' End Type                                           '
''''
''''Private Declare Function SendMessage Lib "user32" _
''''        Alias "SendMessageA" _
''''        (ByVal hWnd As Long, _
''''        ByVal wMsg As Long, _
''''        ByVal wParam As Long, _
''''        lParam As Any) As Long
''''
''''Public Enum eDataType
''''    omString
''''    omDate
''''    omTime
''''    omBoolean
''''    omNumber
''''End Enum
''''
''''Private mCurrentItem As Long
''''Private mDirty As Boolean
''''
''''Private frmParent As PropertyBox
''''
''''Public Event comboListClick(pItem As Long)
''''Public Event ButtonClick(pItem As Long)
''''
''''
''''
''''
''''
''''
''''Public stButtonClick As Boolean
''''
''''
''''
''''Private WithEvents mFG As MSFlexGrid
''''Private WithEvents mTXT As VB.TextBox
''''Private WithEvents mList As VB.ListBox
''''Private WithEvents mCmd As VB.PictureBox
''''Private WithEvents mChk As VB.CheckBox
''''
''''Private stScroll As Boolean
''''
''''Public Sub Init(pFG As MSFlexGrid)
''''    Dim State As Long
''''    Dim i As Integer
''''
''''    'Set frmParent = pFG.Parent
''''
''''
''''
''''    Set mFG = pFG
''''    Set mFG.Parent.Font = mFG.Font
''''
''''
''''    Set mList = mFG.Parent.Controls.Add("VB.ListBox", "cmdList")
''''    mList.ZOrder 0
''''    mList.Appearance = vbFlat
''''    ' mList.Height = 75 '
''''
''''    mList.AddItem "1"
''''    mList.AddItem "2"
''''
''''
''''    Set mTXT = mFG.Parent.Controls.Add("VB.TextBox", "txtValue")
''''
''''    mTXT.ZOrder 0
''''    mTXT.Appearance = vbFlat
''''    mTXT.BorderStyle = 0
''''    mTXT.BackColor = vbWhite
''''    mTXT.Visible = False
''''
''''
''''    Set mCmd = mFG.Parent.Controls.Add("VB.PictureBox", "cmdCombo")
''''    mCmd.ZOrder 0
''''    mCmd.Appearance = vbFlat
''''    mCmd.BorderStyle = 0
'''''    mCmd.Picture = frmParent.ImageList2.ListImages(2).Picture
''''
''''
''''    Set mTXT.Font = mFG.Font
''''    Set mList.Font = mFG.Font
''''
''''
''''
''''    Set mChk = mFG.Parent.Controls.Add("VB.Checkbox", "chkList")
''''    mChk.ZOrder 0
''''    mChk.Caption = ""
''''    mChk.BackColor = vbWhite
''''
''''
''''    For i = 0 To mFG.Rows - 1
''''        If Trim(mFG.TextMatrix(i, 1)) <> "" Then
''''
''''
''''
''''            Set mChk = mFG.Parent.Controls.Add("VB.Checkbox", "chkList" & i)
''''            mChk.ZOrder 0
''''
''''            mChk.Caption = ""
''''            mChk.BackColor = vbWhite
''''
''''
''''
''''            With mChk
''''
''''                mFG.Col = 0
''''                .Top = mFG.Top + mFG.RowPos(i) + 50
''''                .Height = mFG.RowHeight(i) - 10
''''
''''                If mChk.Top + mChk.Height < mFG.Top + mFG.Height Then
''''                    .Visible = True
''''                    .Left = mFG.Left + mFG.ColPos(0) + 90
''''                    .Width = mFG.ColWidth(0) - 200
''''                Else
''''                    .Visible = False
''''                End If
''''
''''            End With
''''
''''
''''        End If
''''    Next
''''End Sub
''''Sub Display()
''''
''''    Const lvBorder = 3
''''    Const lvGrid = 1
''''    Dim pMode As String
''''    Dim pFont As New StdFont
''''    Dim i As Integer
''''    Dim CurrentCellWidth As Long
''''
''''
''''
''''
''''    Static prevTop As Long
''''    Static prevHeight As Long
''''
''''
''''
''''
''''    If mList.Visible = True Then mList.Visible = False
''''
''''
''''    pMode = mFG.TextMatrix(mFG.Row, 4)
''''
''''
''''    Select Case pMode
''''
''''        Case "TEXT"
''''
''''            mCmd.Visible = False
''''
''''            With mTXT
''''
''''                .Visible = True
''''                mFG.Col = 2
''''                If mFG.RowPos(mFG.Row) < 0 Then
''''                    .Visible = False
''''
''''                    Exit Sub
''''
''''                Else
''''                    .Top = mFG.Top + mFG.RowPos(mFG.Row) + 50
''''                    .Height = mFG.RowHeight(mFG.Row) - 10
''''                    .Left = mFG.Left + mFG.ColPos(2) + 120
''''                    prevTop = .Top
''''
''''
''''                    If .Top + .Height < mFG.Top + mFG.Height Then
''''
''''
''''                        .Width = mFG.ColWidth(2) - 100
''''
''''
''''                        '.BackColor = &HFFC0C0
''''                        .BackColor = vbWhite
''''                        .Text = Trim(mFG.Text)
''''                        .Visible = True
''''                    Else
''''                        .Visible = False
''''                    End If
''''
''''
''''                End If
''''
''''                mFG.Col = 1
''''                mFG.HighLight = flexHighlightAlways
''''                .ZOrder 0
''''            End With
''''
''''
''''        Case "BUTTON", "COMBO"
''''            mCmd.Visible = False
''''            If pMode = "BUTTON" Then
''''
'''''              mCmd.Picture = frmParent.ImageList2.ListImages(2).Picture
''''            ElseIf pMode = "COMBO" Then
'''''              mCmd.Picture = frmParent.ImageList2.ListImages(1).Picture
''''            End If
''''
''''            With mTXT
''''
''''                If .Visible Then
''''                    .Visible = True
''''                End If
''''
''''                mFG.Col = 2
''''                If mFG.RowPos(mFG.Row) < 0 Then
''''                    .Visible = False
''''                     mCmd.Visible = False
''''                Else
''''                    .Top = mFG.Top + mFG.RowPos(mFG.Row) + 50
''''                    prevTop = .Top
''''
''''
''''
''''                    .Left = mFG.Left + mFG.ColPos(2) + 120
''''
''''                    .Height = mFG.RowHeight(mFG.Row) - 10
''''
''''                    If .Top + .Height < mFG.Top + mFG.Height Then
''''                        .Visible = True
''''
''''                        .Width = mFG.ColWidth(2) - .Height - 70
''''
''''
''''                        '.BackColor = &HFFC0C0
''''                        .BackColor = vbWhite
''''                        .Text = Trim(mFG.Text)
''''                    Else
''''                        .Visible = False
''''                    End If
''''
''''
''''
''''                End If
''''
''''
''''                mFG.Col = 1
''''                mFG.HighLight = flexHighlightAlways
''''                .ZOrder 0
''''            End With
''''
''''
''''            If mTXT.Visible = True Then
''''                mCmd.Visible = True
''''                mTXT.ZOrder 0
''''                mCmd.ZOrder 0
''''                mCmd.Height = mTXT.Height + 1   ' + 10 '
''''                mCmd.Width = mTXT.Height
''''                mCmd.Top = mTXT.Top
''''                mCmd.Left = mTXT.Left + mTXT.Width
''''                mCmd.ZOrder 0
''''            Else
''''                mCmd.Visible = False
''''            End If
''''
''''
''''    End Select
''''
''''
''''
''''
''''
''''    For i = 0 To mFG.Rows - 1
''''
''''        Debug.Print i & " : " & mFG.RowPos(i)
''''
''''        Err.Clear
''''        On Error Resume Next
''''        Set mChk = mFG.Parent.Controls("chkList" & i)
''''        If Err.Number = 0 Then
''''            mChk.Visible = False
''''            If mFG.RowPos(i) < 0 Then
''''                mChk.Visible = False
''''                Debug.Print mChk.Name
''''            Else
''''
''''                Debug.Print mChk.Name
''''                mChk.Left = mFG.Left + mFG.ColPos(0) + 90
''''
''''                mChk.Height = mFG.RowHeight(i) - 10
''''                mFG.Col = 0
''''                mChk.Width = 0
''''                mChk.Width = mFG.ColWidth(0) - 200
''''                mChk.Top = mFG.Top + mFG.RowPos(i) + 50
''''
''''                If mChk.Top + mChk.Height < mFG.Top + mFG.Height Then
''''                    mChk.Visible = True
''''                Else
''''                    mChk.Visible = False
''''                End If
''''
''''            End If
''''
''''
''''        End If
''''    Next
''''
''''
''''
''''    mFG.Col = 1
''''    mFG.HighLight = flexHighlightAlways
''''
''''
''''End Sub
''''
''''Private Sub Class_Terminate()
''''    Set mTXT = Nothing
''''    Set mCmd = Nothing
''''    Set mList = Nothing
''''End Sub
''''
''''Private Sub mCmd_Click()
''''    Dim tmp As Variant
''''    Dim pMode As String
''''
''''
''''
''''
''''    pMode = mFG.TextMatrix(mFG.Row, 4)
''''
''''
''''    Select Case pMode
''''        Case "COMBO"
''''            stButtonClick = False
''''
''''            If mList.Visible = True Then
''''
''''                mList.Visible = False
''''
''''            Else
''''
''''                mList.Top = mTXT.Top + mTXT.Height
''''                mList.Width = mTXT.Width + mCmd.Width
''''                mList.Left = mTXT.Left
''''
''''                tmp = mFG.Text
''''                mList.Clear
''''
''''                mList.AddItem "1"
''''                mList.AddItem "2"
''''                mList.AddItem "3"
''''
''''                mList.Visible = True
''''                mList.ZOrder 0
''''
''''            End If
''''        Case "TEXT"
''''            stButtonClick = False
''''        Case "BUTTON"
''''            RaiseEvent ButtonClick(mFG.Row)
''''            stButtonClick = True
''''
''''    End Select
''''End Sub
''''
''''
''''
''''Private Sub mFG_EnterCell()
''''    mFG.HighLight = flexHighlightNever
''''End Sub
''''
''''Private Sub mFG_Scroll()
''''   Display
''''
''''End Sub
''''
''''Private Sub mList_Click()
''''
''''
''''    mTXT = mList.Text
''''
''''    mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
''''''    mFG.Text = mTXT
''''    RaiseEvent comboListClick(mFG.Row)
''''
''''
''''End Sub
''''
''''Private Sub mList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''
''''    mList.Visible = False
''''
''''End Sub
''''
''''
''''
''''Private Sub mFG_Click()
''''    Dim stLock As Boolean
''''
''''
''''    mTXT.Visible = True
''''    mTXT.ZOrder 0
''''
'''''''    If mDirty Then
'''''''''''        mFG.Text = mTXT.Text
'''''''
'''''''        mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
'''''''    End If
''''
''''
''''    mTXT.Text = mFG.TextMatrix(mFG.Row, 2)
''''
''''    mCurrentItem = mFG.Row
''''
''''    mCmd.Visible = True
''''    Display
''''
''''
''''
''''    mCmd.ZOrder 0
''''    mTXT.ZOrder 0
''''
''''End Sub
''''Private Sub mFG_KeyPress(KeyAscii As Integer)
''''
''''    mTXT.Text = Chr(KeyAscii)
''''    mTXT.SelStart = 1
''''
''''    If mTXT.Visible Then
''''       mTXT.SetFocus
''''    Else
''''      mTXT.Visible = True
''''      mTXT.SetFocus
''''
''''    End If
''''
''''End Sub
''''
''''Private Sub mTXT_Change()
''''
''''''''    mFG.Text = " " & Trim(CStr(mTXT.Text))
''''    mDirty = True
''''    mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
''''End Sub
''''
''''Private Sub mTXT_KeyPress(KeyAscii As Integer)
''''
''''    If KeyAscii = 13 Then
''''
''''''''       mFG.Text = mTXT.Text
''''       mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
''''    End If
''''
''''End Sub
''''Public Sub Resize()
''''
''''    mTXT.Visible = False
''''    mCmd.Visible = False
''''
''''    mFG.Width = mFG.Parent.ScaleWidth - mFG.Left
''''
''''    ' mFG.ColumnHeaders(2).Width = mFG.Parent.ScaleWidth - mFG.ColumnHeaders(1).Width - 50 '
''''    mFG.Height = mFG.Parent.Height
''''
''''End Sub
''''
''''
''''
''''
''''Private Function getTextWidth(pString As String) As Long
''''
''''    getTextWidth = mFG.Parent.TextWidth(pString)
''''
''''End Function
''''
''''Public Sub mTXT_LostFocus()
''''' mFG.Col = 2          '
''''' mFG.Text = mTXT.Text '
'''''mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
''''End Sub
''''
''''
''''Sub AddProperty(grid As MSFlexGrid, strProperty As String, Optional tag1 As String, Optional tag2 As String)
''''
''''  If grid.Rows = 1 Then
''''    If Trim(grid.TextMatrix(0, 1)) = "" Then
''''      grid.TextMatrix(0, 1) = strProperty
''''      grid.TextMatrix(0, 4) = tag1
''''      grid.TextMatrix(0, 5) = tag2
''''    Else
''''      grid.Rows = grid.Rows + 1
''''      grid.TextMatrix(grid.Rows - 1, 1) = strProperty
''''      grid.TextMatrix(grid.Rows - 1, 4) = tag1
''''      grid.TextMatrix(grid.Rows - 1, 5) = tag2
''''    End If
''''
''''  Else
''''    grid.Rows = grid.Rows + 1
''''    grid.TextMatrix(grid.Rows - 1, 1) = strProperty
''''    grid.TextMatrix(grid.Rows - 1, 4) = tag1
''''    grid.TextMatrix(grid.Rows - 1, 5) = tag2
''''  End If
''''End Sub
''''
