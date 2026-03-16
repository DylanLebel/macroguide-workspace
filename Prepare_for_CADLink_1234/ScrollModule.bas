Attribute VB_Name = "ScrollModule"
Option Explicit
Option Private Module

#If VBA7 Then
    Public Declare PtrSafe Function SetWindowSubclass Lib "comctl32" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Boolean
    Public Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32" (ByVal hwnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Boolean
    Public Declare PtrSafe Function DefSubclassProc Lib "comctl32" (ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
#Else
    Public Declare Function SetWindowSubclass Lib "comctl32" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Boolean
    Public Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Boolean
    Public Declare Function DefSubclassProc Lib "comctl32" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_VSCROLL As Long = &H115
Private Const SB_LINEUP As Long = 0
Private Const SB_LINEDOWN As Long = 1

#If VBA7 Then
Public Function ListBoxWndProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
#Else
Public Function ListBoxWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
#End If
    If uMsg = WM_MOUSEWHEEL Then
        Dim zDelta As Integer
        zDelta = HiWord(wParam)
        
        If zDelta > 0 Then
            SendMessage hwnd, WM_VSCROLL, SB_LINEUP, 0
        Else
            SendMessage hwnd, WM_VSCROLL, SB_LINEDOWN, 0
        End If
        
        ListBoxWndProc = 0 ' Prevent default processing of this message
    Else
        ListBoxWndProc = DefSubclassProc(hwnd, uMsg, wParam, lParam)
    End If
End Function

Private Function HiWord(ByVal dwValue As Long) As Integer
    HiWord = (dwValue And &H8000&) * &H10000 Or (dwValue And &H7FFF&)
End Function

