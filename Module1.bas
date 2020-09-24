Attribute VB_Name = "Module1"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
' You can find more o these (lower) in the API Viewer.  Here
' they are used only for resizing the left and right
Public Const HTLEFT = 10
Public Const HTRIGHT = 11

