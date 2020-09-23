Attribute VB_Name = "Module1"
Option Explicit

Public rs As DAO.Recordset
Public dbs As DAO.Database

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Long, _
ByVal bErase As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Public Const GW_MAX = 5
Public Const WS_EX_LAYOUTRTL = &H400000
Public Const GWL_EXSTYLE = (-20)
Public Sub ConnectAccessDb()
    
    Set dbs = OpenDatabase(App.Path & "\Test.mdb", False, False, ";Pwd=admin???")
    
End Sub
Public Sub CloseAccessDb()
  
    rs.Close
    dbs.Close
    Set rs = Nothing
    Set dbs = Nothing

End Sub

