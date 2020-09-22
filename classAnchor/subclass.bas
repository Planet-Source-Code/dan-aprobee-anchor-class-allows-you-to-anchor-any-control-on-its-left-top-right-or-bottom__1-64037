Attribute VB_Name = "subclass"
Option Explicit

Private defWindowProc As Long
Public minX As Long
Public minY As Long
Public maxX As Long
Public maxY As Long

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_GETMINMAXINFO As Long = &H24

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)


Public Sub SubClass(hwnd As Long)

   On Error Resume Next
   defWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, _
                        AddressOf WindowProc)
End Sub


Public Sub UnSubClass(hwnd As Long)

  'restore the default message handling
  'before exiting
   If defWindowProc Then
      SetWindowLong hwnd, GWL_WNDPROC, defWindowProc
      defWindowProc = 0
   End If
   
End Sub


Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
   On Error Resume Next
 
   Select Case uMsg
      Case WM_GETMINMAXINFO
         Dim MMI As MINMAXINFO
         CopyMemory MMI, ByVal lParam, LenB(MMI)
         
            With MMI
              .ptMinTrackSize.x = minX
              .ptMinTrackSize.y = minY
              .ptMaxTrackSize.x = maxX
              .ptMaxTrackSize.y = maxY
            End With
      
            CopyMemory ByVal lParam, MMI, LenB(MMI)
            WindowProc = 0
       Case Else
          WindowProc = CallWindowProc(defWindowProc, _
                      hwnd, uMsg, wParam, lParam)
     End Select
   
End Function


