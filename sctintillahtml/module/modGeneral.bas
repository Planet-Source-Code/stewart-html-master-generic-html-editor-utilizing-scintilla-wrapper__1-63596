Attribute VB_Name = "modGeneral"
Private Type tagInitCommonControlsEx
lngSize As Long
lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
(iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
On Error Resume Next
Dim iccex As tagInitCommonControlsEx
' Ensure CC available:
With iccex
.lngSize = LenB(iccex)
.lngICC = ICC_USEREX_CLASSES
End With
InitCommonControls
InitCommonControlsEx iccex

InitCommonControlsVB = (Err.Number = 0)
  
End Function

Public Sub Main()
frmMain.Show
End Sub


' Thanks to Charlie Wilson for this function
Public Function HexRGB(lCdlColor As Long)

    Dim lCol As Long
    Dim iRed, iGreen, iBlue As Integer
    Dim vHexR, vHexG, vHexB As Variant
    'Break out the R, G, B values from the c
    '     ommon dialog color
    lCol = lCdlColor
    iRed = lCol Mod &H100
    lCol = lCol \ &H100
    iGreen = lCol Mod &H100
    lCol = lCol \ &H100
    iBlue = lCol Mod &H100
    
    'Determine Red Hex
    vHexR = Hex(iRed)


    If Len(vHexR) < 2 Then
        vHexR = "0" & vHexR
    End If

    'Determine Green Hex
    vHexG = Hex(iGreen)


    If Len(vHexG) < 2 Then
        vHexG = "0" & iGreen
    End If

    'Determine Blue Hex
    vHexB = Hex(iBlue)


    If Len(vHexB) < 2 Then
        vHexB = "0" & vHexB
    End If

    'Add it up, return the function value
    HexRGB = "#" & vHexR & vHexG & vHexB
End Function

Public Function ExtractFileName(ByVal strPath As String) As String

    ' StrReverse is only working in VB6
    strPath = StrReverse(strPath)
    strPath = Left(strPath, InStr(strPath, "\") - 1)
    ExtractFileName = StrReverse(strPath)
End Function





