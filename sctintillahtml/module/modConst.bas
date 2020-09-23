Attribute VB_Name = "modConst"
Option Explicit

Public Const WM_NOTIFY = &H4E

Public Const PHYSICALWIDTH = 110 '  Physical Width in device units
Public Const PHYSICALHEIGHT = 111 '  Physical Height in device units

Public Const strQ = """"

Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Public Type SCNotification
    NotifyHeader As NMHDR
    position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    length As Long
    linesAdded As Long
    Message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    x As Long
    y As Long
End Type

Public Enum EOL
    SC_EOL_CRLF = 0                     ' CR + LF
    SC_EOL_CR = 1                       ' CR
    SC_EOL_LF = 2                       ' LF
End Enum


Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)

    
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer


    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With

    

    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If

    Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub
