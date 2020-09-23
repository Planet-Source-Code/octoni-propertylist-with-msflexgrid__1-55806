VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl PropertyBox 
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   ScaleHeight     =   5430
   ScaleWidth      =   7830
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertyBox.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertyBox.ctx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertyBox.ctx":0644
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertyBox.ctx":0A0A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   -2147483638
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
   End
End
Attribute VB_Name = "PropertyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Locked = 0
Const m_def_CheckboxVisible = 0
Const m_def_PropertyCount = 0
'Property Variables:
Dim m_Locked As Variant
Dim m_CheckboxVisible As Variant
Dim m_PropertyCount As Variant
'Event Declarations:
Event ButtonClick(strProperty As String)
Event ComboClick(strProperty As String)





Private Const LVM_FIRST As Long = &H1000
' Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)        '
' Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57) '
' Private Const LVHT_NOWHERE As Long = &H1                    '
' Private Const LVHT_ONITEMICON As Long = &H2                 '
' Private Const LVHT_ONITEMLABEL As Long = &H4                '
' Private Const LVHT_ONITEMSTATEICON As Long = &H8            '


Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
' Private Const LVS_EX_SUBITEMIMAGES As Long = &H2                      '
' Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54) '
' Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55) '


' Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or _
LVHT_ONITEMLABEL Or _
        LVHT_ONITEMSTATEICON)

Private Const LVIR_BOUNDS = 0
Private Const LVIR_ICON = 1
Private Const LVIR_LABEL = 2
Private Const LVIR_SELECTBOUNDS = 3



Private Const LVMODE_TEXT = 0
Private Const LVMODE_COMBO = 1
Private Const LVMODE_BUTTON = 2


' Private Type POINTAPI '
' X As Long             '
' Y As Long             '
' End Type              '

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Private Type LVHITTESTINFO                         '
' pt As POINTAPI                                     '
' flags As Long                                      '
' iItem As Long                                      '
' iSubItem  As Long  'ie3+ only .. was NOT in win95. '
' 'Valid only for LVM_SUBITEMHITTEST                 '
' End Type                                           '

Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" _
        (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long

Private Enum eDataType
    omString
    omDate
    omTime
    omBoolean
    omNumber
End Enum

Private mCurrentItem As Long
Private mDirty As Boolean

Private frmParent As PropertyBox



Private stButtonClick As Boolean



Private WithEvents mFG As MSFlexGrid
Attribute mFG.VB_VarHelpID = -1
Private WithEvents mTXT As VB.TextBox
Attribute mTXT.VB_VarHelpID = -1
Private WithEvents mList As VB.ListBox
Attribute mList.VB_VarHelpID = -1
Private WithEvents mCmd As VB.PictureBox
Attribute mCmd.VB_VarHelpID = -1
Private WithEvents mChk As VB.CheckBox
Attribute mChk.VB_VarHelpID = -1

Private stScroll As Boolean

Private stCheckBoxVisible As Boolean
Private stLocked As Boolean

Private Sub HideCheckBox()
    Dim ctl As Control
    
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = MSFlexGrid1.Width / 2
    MSFlexGrid1.ColWidth(2) = MSFlexGrid1.Width / 2
    stCheckBoxVisible = False
    
    
    For Each ctl In UserControl.Controls
      If TypeOf ctl Is CheckBox Then
        ctl.Visible = False
      End If
    Next
End Sub

Private Sub ShowCheckBox()
 

   stCheckBoxVisible = True
   MSFlexGrid1.ColWidth(0) = 400
   MSFlexGrid1.ColWidth(1) = MSFlexGrid1.Width / 2 - 500
   MSFlexGrid1.ColWidth(2) = MSFlexGrid1.Width - (MSFlexGrid1.ColWidth(0) + MSFlexGrid1.ColWidth(1)) - 350

    Display
End Sub
Public Sub DrawPropertyBox()
    Dim State As Long
    Dim i As Integer


    
     
    MSFlexGrid1.ColWidth(0) = 400
    MSFlexGrid1.ColWidth(1) = MSFlexGrid1.Width / 2 - 500
    MSFlexGrid1.ColWidth(2) = MSFlexGrid1.Width - (MSFlexGrid1.ColWidth(0) + MSFlexGrid1.ColWidth(1)) - 350

    stCheckBoxVisible = True
    stLocked = False

    



    MSFlexGrid1.ScrollTrack = True
    
    Set mFG = MSFlexGrid1

    
    Set mList = UserControl.Controls.Add("VB.ListBox", "cmdList")
    mList.ZOrder 0
    mList.Appearance = vbFlat
    ' mList.Height = 75 '



    Set mTXT = UserControl.Controls.Add("VB.TextBox", "txtValue")

    mTXT.ZOrder 0
    mTXT.Appearance = vbFlat
    mTXT.BorderStyle = 0
    mTXT.BackColor = vbWhite
    mTXT.Visible = False


    Set mCmd = UserControl.Controls.Add("VB.PictureBox", "cmdCombo")
    mCmd.ZOrder 0
    mCmd.Appearance = vbFlat
    mCmd.BorderStyle = 0
'    mCmd.Picture = frmParent.ImageList2.ListImages(2).Picture


    Set mTXT.Font = mFG.Font
    Set mList.Font = mFG.Font



    Set mChk = UserControl.Controls.Add("VB.Checkbox", "chkList")
    mChk.ZOrder 0
    mChk.Caption = ""
    mChk.BackColor = vbWhite


    For i = 0 To mFG.Rows - 1
        If Trim(mFG.TextMatrix(i, 1)) <> "" Then



            Set mChk = UserControl.Controls.Add("VB.Checkbox", "chkList" & i)
            mChk.ZOrder 0

            mChk.Caption = ""
            mChk.BackColor = vbWhite



            With mChk

                mFG.Col = 0
                .Top = mFG.Top + mFG.RowPos(i) + 50
                .Height = mFG.RowHeight(i) - 10

                If mChk.Top + mChk.Height < mFG.Top + mFG.Height Then
                    .Visible = True
                    .Left = mFG.Left + mFG.ColPos(0) + 90
                    .Width = mFG.ColWidth(0) - 200
                Else
                    .Visible = False
                End If

            End With


        End If
    Next
End Sub
Sub Display()

    Const lvBorder = 3
    Const lvGrid = 1
    Dim pMode As String
    Dim pFont As New StdFont
    Dim i As Integer
    Dim CurrentCellWidth As Long




    Static prevTop As Long
    Static prevHeight As Long




    If mList.Visible = True Then mList.Visible = False
    
    
    pMode = mFG.TextMatrix(mFG.Row, 4)
    

    Select Case pMode

        Case "TEXT"

            mCmd.Visible = False

            With mTXT

                .Visible = True
                mFG.Col = 2
                If mFG.RowPos(mFG.Row) < 0 Then
                    .Visible = False

                    Exit Sub

                Else
                    .Top = mFG.Top + mFG.RowPos(mFG.Row) + 50
                    .Height = mFG.RowHeight(mFG.Row) - 10
                    
                    
                    
                    .Left = mFG.Left + mFG.ColPos(2) + 120
                    
                    
                    
                    prevTop = .Top
                    
                    
                    If .Top + .Height < mFG.Top + mFG.Height Then
                        
                        If stCheckBoxVisible Then
                          .Width = mFG.ColWidth(2) - 100
                        Else
                          .Width = mFG.ColWidth(2) - 420
                        End If

                        

                        '.BackColor = &HFFC0C0
                        .BackColor = vbWhite
                        .Text = Trim(mFG.Text)
                        .Visible = True
                    Else
                        .Visible = False
                    End If
                    
                    
                End If
     
                mFG.Col = 1
                mFG.HighLight = flexHighlightAlways
                .ZOrder 0
            End With


        Case "BUTTON", "COMBO"
            mCmd.Visible = False
            If pMode = "BUTTON" Then
            
              mCmd.Picture = ImageList2.ListImages(2).Picture
            ElseIf pMode = "COMBO" Then
              mCmd.Picture = ImageList2.ListImages(1).Picture
            End If

            With mTXT

                If .Visible Then
                    .Visible = True
                End If

                mFG.Col = 2
                If mFG.RowPos(mFG.Row) < 0 Then
                    .Visible = False
                     mCmd.Visible = False
                Else
                    .Top = mFG.Top + mFG.RowPos(mFG.Row) + 50
                    prevTop = .Top

                   .Left = mFG.Left + mFG.ColPos(2) + 120
                 
                   .Height = mFG.RowHeight(mFG.Row) - 10

                    If .Top + .Height < mFG.Top + mFG.Height Then
                        .Visible = True

                        If stCheckBoxVisible Then
                          .Width = mFG.ColWidth(2) - .Height - 70
                        Else
                          .Width = mFG.ColWidth(2) - .Height - 430
                        End If


                        '.BackColor = &HFFC0C0
                        .BackColor = vbWhite
                        .Text = Trim(mFG.Text)
                    Else
                        .Visible = False
                    End If



                End If


                mFG.Col = 1
                mFG.HighLight = flexHighlightAlways
                .ZOrder 0
            End With


            If mTXT.Visible = True Then
                mCmd.Visible = True
                mTXT.ZOrder 0
                mCmd.ZOrder 0
                mCmd.Height = mTXT.Height + 1   ' + 10 '
                mCmd.Width = mTXT.Height
                mCmd.Top = mTXT.Top
                
                mCmd.Left = mTXT.Left + mTXT.Width
                
                
                mCmd.ZOrder 0
            Else
                mCmd.Visible = False
            End If


    End Select


   
   
    If stCheckBoxVisible = True Then


    For i = 0 To mFG.Rows - 1

        'debug.print i & " : " & mFG.RowPos(i)

        err.Clear
        On Error Resume Next
        Set mChk = UserControl.Controls("chkList" & i)
        If err.Number = 0 Then
            mChk.Visible = False
            If mFG.RowPos(i) < 0 Then
                mChk.Visible = False
                'debug.print mChk.Name
            Else

                'debug.print mChk.Name
                mChk.Left = mFG.Left + mFG.ColPos(0) + 90

                mChk.Height = mFG.RowHeight(i) - 10
                mFG.Col = 0
                mChk.Width = 0
                mChk.Width = mFG.ColWidth(0) - 200
                mChk.Top = mFG.Top + mFG.RowPos(i) + 50

                If mChk.Top + mChk.Height < mFG.Top + mFG.Height Then
                    mChk.Visible = True
                Else
                    mChk.Visible = False
                End If

            End If


        End If
    Next

    End If

    mFG.Col = 1
    mFG.HighLight = flexHighlightAlways


End Sub

Private Sub Class_Terminate()
    Set mTXT = Nothing
    Set mCmd = Nothing
    Set mList = Nothing
End Sub

Private Sub mCmd_Click()
    Dim tmp As String
    Dim pMode As String
    Dim i As Long
    Dim objMatch As Match
    Dim objMatches As MatchCollection
    Dim regex As New clsRegEx

    
    
    If stLocked = True Then
      Exit Sub
    End If
    

    pMode = mFG.TextMatrix(mFG.Row, 4)


    Select Case pMode
        Case "COMBO"
            stButtonClick = False

            If mList.Visible = True Then

                mList.Visible = False

            Else

                mList.Top = mTXT.Top + mTXT.Height
                mList.Width = mTXT.Width + mCmd.Width
                mList.Left = mTXT.Left

                tmp = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
                mList.Clear
                
                If regex.MatchRegex("\|", tmp) Then
                  Set objMatches = regex.GetMatchCollection("\|", tmp)
                  For i = 0 To objMatches.count - 1
                    If i = objMatches.count - 1 Then
                      Exit For
                    End If
                    mList.AddItem Mid(tmp, objMatches.Item(i).FirstIndex + 2, objMatches.Item(i + 1).FirstIndex - objMatches.Item(i).FirstIndex - 1)
                    DoEvents
                  Next
                End If
                  
'                For i = 0 To UBound(tmp)
'
'                Next

                

                mList.Visible = True
                mList.ZOrder 0

            End If
        Case "TEXT"
            stButtonClick = False
        Case "BUTTON"
            RaiseEvent ButtonClick(mFG.TextMatrix(mFG.Row, 1))
            stButtonClick = True

    End Select
End Sub



Private Sub mFG_EnterCell()
    mFG.HighLight = flexHighlightNever
End Sub

Private Sub mFG_Scroll()
   Display

End Sub

Private Sub mList_Click()


    mTXT = mList.Text
    
    mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
''    mFG.Text = mTXT

    RaiseEvent ComboClick(mFG.Row)


End Sub

Private Sub mList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mList.Visible = False

End Sub



Private Sub mFG_Click()



    mTXT.Visible = True
    mTXT.ZOrder 0

'''    If mDirty Then
'''''''        mFG.Text = mTXT.Text
'''
'''        mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
'''    End If


    mTXT.Text = mFG.TextMatrix(mFG.Row, 2)

    mCurrentItem = mFG.Row

    mCmd.Visible = True
    Display



    mCmd.ZOrder 0
    mTXT.ZOrder 0

End Sub
Private Sub mFG_KeyPress(KeyAscii As Integer)

    mTXT.Text = Chr(KeyAscii)
    mTXT.SelStart = 1

    If mTXT.Visible Then
       mTXT.SetFocus
    Else
      mTXT.Visible = True
      mTXT.SetFocus

    End If

End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  If stLocked Then
    KeyCode = 0
  End If
  
  
  
  
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
  If stLocked Then
   KeyAscii = 0
  End If
  
End Sub

Private Sub mTXT_Change()

''''    mFG.Text = " " & Trim(CStr(mTXT.Text))
    If stLocked Then
      Exit Sub
    End If
    
    mDirty = True
    mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
End Sub

Private Sub mTXT_KeyDown(KeyCode As Integer, Shift As Integer)
    If stLocked = True Then
      KeyCode = 0
      Exit Sub
    End If
    
End Sub

Private Sub mTXT_KeyPress(KeyAscii As Integer)
    
    
    If stLocked = True Then
      KeyAscii = 0
      Exit Sub
    End If
    
    
    If KeyAscii = 13 Then

''''       mFG.Text = mTXT.Text
       mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
    End If

End Sub
Private Sub Resize()

    mTXT.Visible = False
    mCmd.Visible = False

    mFG.Width = UserControl.ScaleWidth - mFG.Left

    ' mFG.ColumnHeaders(2).Width =UserControl.ScaleWidth - mFG.ColumnHeaders(1).Width - 50 '
    mFG.Height = UserControl.Height

End Sub




Private Function getTextWidth(pString As String) As Long

    getTextWidth = UserControl.TextWidth(pString)

End Function

Private Sub mTXT_LostFocus()
' mFG.Col = 2          '
' mFG.Text = mTXT.Text '
'mFG.TextMatrix(mFG.Row, 2) = " " & Trim(CStr(mTXT.Text))
End Sub

























'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PropertyCount() As Variant
    PropertyCount = m_PropertyCount
End Property

Private Property Let PropertyCount(ByVal New_PropertyCount As Variant)
    m_PropertyCount = New_PropertyCount
    PropertyChanged "PropertyCount"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetPropertyValue(strProperty As String) As Variant
  Dim i As Long
  
  For i = 0 To MSFlexGrid1.Rows - 1
    If UCase(MSFlexGrid1.TextMatrix(i, 1)) = UCase(Trim(strProperty)) Then
      GetPropertyValue = MSFlexGrid1.TextMatrix(i, 2)
      Exit Function
    End If
  Next
  GetPropertyValue = ""


End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function AddProperty(strProperty As String, Optional PropertyType As String, Optional Tag As String, Optional strListValues As String) As Variant
   

   
   
   If MSFlexGrid1.Rows = 1 Then
    If Trim(MSFlexGrid1.TextMatrix(0, 1)) = "" Then
      MSFlexGrid1.TextMatrix(0, 1) = strProperty
      MSFlexGrid1.TextMatrix(0, 4) = PropertyType
      MSFlexGrid1.TextMatrix(0, 5) = Tag
      MSFlexGrid1.TextMatrix(0, 6) = strListValues
    Else
      MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = strProperty
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = PropertyType
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = Tag
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 6) = strListValues
    End If
    
  Else
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = strProperty
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = PropertyType
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = Tag
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 6) = strListValues
  End If
  
  PropertyCount = MSFlexGrid1.Rows
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function RemoveProperty(index As Long) As Variant
  MSFlexGrid1.RemoveItem index
End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

   
 
    m_PropertyCount = m_def_PropertyCount
    
   
    
    m_CheckboxVisible = m_def_CheckboxVisible
    m_Locked = m_def_Locked
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_PropertyCount = PropBag.ReadProperty("PropertyCount", m_def_PropertyCount)
    m_CheckboxVisible = PropBag.ReadProperty("CheckboxVisible", m_def_CheckboxVisible)
    m_Locked = PropBag.ReadProperty("Locked", m_def_Locked)
End Sub

Private Sub UserControl_Resize()
  With MSFlexGrid1
    .Top = 0
    .Left = 0
    .Width = UserControl.Width
    .Height = UserControl.Height
  End With
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("PropertyCount", m_PropertyCount, m_def_PropertyCount)
    Call PropBag.WriteProperty("CheckboxVisible", m_CheckboxVisible, m_def_CheckboxVisible)
    Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetTag(strProperty As String) As Variant
  Dim i As Long
  
  For i = 0 To MSFlexGrid1.Rows - 1
    If UCase(MSFlexGrid1.TextMatrix(i, 1)) = UCase(Trim(strProperty)) Then
      GetTag = MSFlexGrid1.TextMatrix(i, 5)
      Exit Function
    End If
  Next
  GetTag = ""
  
End Function



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetCheckStatus(strProperty As String) As Variant
  Dim i As Long
  
  For i = 0 To MSFlexGrid1.Rows - 1
    If UCase(MSFlexGrid1.TextMatrix(i, 1)) = UCase(Trim(strProperty)) Then
      Set mChk = UserControl.Controls("chkList" & i)
      GetCheckStatus = mChk.Value
      Exit Function
    End If
  Next
  GetCheckStatus = 0
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CheckboxVisible() As Variant
    CheckboxVisible = m_CheckboxVisible
End Property

Public Property Let CheckboxVisible(ByVal New_CheckboxVisible As Variant)
    m_CheckboxVisible = New_CheckboxVisible
    PropertyChanged "CheckboxVisible"
    
    If New_CheckboxVisible = False Then
      HideCheckBox
    ElseIf New_CheckboxVisible = True Then
      ShowCheckBox
    End If
    
    
    
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Locked() As Variant
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Variant)
    m_Locked = New_Locked
    PropertyChanged "Locked"
    
    
    stLocked = New_Locked
    mTXT.Locked = stLocked
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function SetPropertyValue(strProperty As String, strValue As String) As Variant

 Dim i As Long
  
  For i = 0 To MSFlexGrid1.Rows - 1
    If UCase(MSFlexGrid1.TextMatrix(i, 1)) = UCase(Trim(strProperty)) Then
      mFG.Row = i
      mTXT.Text = strValue
      MSFlexGrid1.TextMatrix(i, 2) = " " & strValue
      mTXT.Refresh
      MSFlexGrid1.Refresh
      Exit Function
    End If
  Next
  
  
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetPropertyName(index As Long) As Variant
  If index <= mFG.Rows - 1 Then
    GetPropertyName = mFG.TextMatrix(index, 1)
  Else
    GetPropertyName = ""
  End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MSFlexGrid1,MSFlexGrid1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    MSFlexGrid1.Refresh
End Sub

