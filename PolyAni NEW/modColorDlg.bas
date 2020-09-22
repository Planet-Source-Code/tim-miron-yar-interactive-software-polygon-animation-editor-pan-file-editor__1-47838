Attribute VB_Name = "modColorDlg"
Private Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Dim CustomColors() As Byte

Public Function ShowColor(ParentForm As Form) As Long
On Error GoTo ErrOut:
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    'set the structure size
    cc.lStructSize = Len(cc)
    
    'Set the owner
    cc.hWndOwner = ParentForm.hWnd
    
    'set the application's instance
    cc.hInstance = App.hInstance
    
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    
    'no extra flags
    cc.flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
    
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If

Exit Function
ErrOut:
  MsgBox Err.Description
End Function

Public Sub PrepColorDlg()
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    Dim I As Integer
    For I = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(I) = 255
    Next I
End Sub
