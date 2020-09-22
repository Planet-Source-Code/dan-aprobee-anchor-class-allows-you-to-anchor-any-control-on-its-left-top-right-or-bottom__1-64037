VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_anchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum enAP
   apTop = 2
   apLeft = 4
   apBottom = 8
   apRight = 16
   apAll = 30
End Enum

Dim mCtrl()      As Object
Dim mCtrlVals()  As Integer
Attribute mCtrlVals.VB_VarHelpID = -1
Dim mctrlCounter As Integer
Dim mForm        As Form

Sub SetFormDimensions(yourForm As Form, _
                      formMinTwipWidth As Long, _
                      formMaxTwipWidth As Long, _
                      formMinTwipHeight As Long, _
                      formMaxTwipHeight As Long)

   Set mForm = yourForm
   'conver the forms min and max vals to pixels
   SubClass.minX = twipToPixX(formMinTwipWidth)
   SubClass.maxX = twipToPixX(formMaxTwipWidth)
   SubClass.minY = twipToPixY(formMinTwipHeight)
   SubClass.maxY = twipToPixY(formMaxTwipHeight)
   'start the subclassing looking for form resizing
   SubClass.SubClass mForm.hwnd
End Sub

Private Function twipToPixX(twipVal As Long) As Long
  twipToPixX = (twipVal / Screen.TwipsPerPixelX)
End Function
Private Function twipToPixY(twipVal As Long) As Long
  twipToPixY = (twipVal / Screen.TwipsPerPixelY)
End Function

Sub Anchor(Control As Object, anchorPoints As enAP)
Dim pLeft    As Long
Dim pRight   As Long
Dim pTop     As Long
Dim pBottom  As Long
Dim pWid     As Long
Dim pHei     As Long

  ReDim Preserve mCtrl(mctrlCounter)
  ReDim Preserve mCtrlVals(6, mctrlCounter)
  
  With Control
    'store control
    Set mCtrl(mctrlCounter) = Control
    'store controls anchorpoint
    mCtrlVals(0, mctrlCounter) = anchorPoints
    'store distance between controls
    'right edge and its parents right edge
    pRight = (.Parent.Width - (.Left + .Width))
    mCtrlVals(1, mctrlCounter) = pRight
    'store distance between controls bottom
    'edge and its parents bottom edge
    pBottom = (.Parent.Height - (.Top + .Height))
    mCtrlVals(2, mctrlCounter) = pBottom
    'store distance between controls left
    'edge and its parents left edge
    pLeft = (.Left)
    mCtrlVals(3, mctrlCounter) = pLeft
    'store distance between controls top
    'edge and the parents top edge
    pTop = .Top
    mCtrlVals(4, mctrlCounter) = pTop
    'store the parents width in case
    'left is not anchored
    pWid = .Parent.Width
    mCtrlVals(5, mctrlCounter) = pWid
    'store the parents height in case
    'top is not anchored
    pHei = .Parent.Height
    mCtrlVals(6, mctrlCounter) = pHei
    'increment the controls counter
    mctrlCounter = (mctrlCounter + 1)
  End With
  
End Sub

Sub Parent_Resize_Event()
Dim ctl               As Object
Dim i                 As Integer
Dim iRight            As Integer
Dim iBottom           As Integer
Dim iLeft             As Integer
Dim iTop              As Integer
Dim iWid              As Integer
Dim iHei              As Integer
Dim bchangeRight      As Boolean
Dim bchangeBottom     As Boolean
Dim bchangeLeft       As Boolean
Dim bchangeTop        As Boolean
Dim ichangeval        As Integer
Dim ianchorPoints     As Integer
Dim imultiplyerLeft   As Single
Dim imultiplyerTop    As Single

  'loop through all controls in array
  For i = 0 To (mctrlCounter - 1)
     'store reference to the control,
     Set ctl = mCtrl(i)
  
     With ctl
       'store controls anchorpoint
       ianchorPoints = mCtrlVals(0, i)
       'store controls distance from
       'right edge of parent
       iRight = mCtrlVals(1, i)
       'store controls distance from
       'bottom edge of parent
       iBottom = mCtrlVals(2, i)
       'store the controls original left
       iLeft = mCtrlVals(3, i)
       'store the parents original width
       iWid = mCtrlVals(5, i)
       'store the controls original top
       iTop = mCtrlVals(4, i)
       'store the parents original height
       iHei = mCtrlVals(6, i)
      
       'reinitialize [bchange*] variables
       bchangeLeft = False
       bchangeTop = False
       bchangeRight = False
       bchangeBottom = False
       'in most instances (except if
       'anchorpoints= 2 or 4)...
       imultiplyerLeft = 1
       imultiplyerTop = 1
       
       Select Case ianchorPoints
          Case Is = 2    'top
              bchangeLeft = True
              imultiplyerLeft = 0.5
          Case Is = 4    'left
              bchangeTop = True
              imultiplyerTop = 0.6
          Case Is = 8    'bottom
              bchangeBottom = True
              bchangeLeft = True
              bchangeTop = True
          Case Is = 6    'top and left
             'since anchoring top and left
             'is default behavior, do nothing
          Case Is = 10   'top and bottom
              bchangeBottom = True
              bchangeLeft = True
          Case Is = 12   'left and bottom
              bchangeBottom = True
              bchangeTop = True
          Case Is = 14   'left and top and bottom
              bchangeBottom = True
          Case Is = 16   'right
              bchangeRight = True
              bchangeLeft = True
              bchangeTop = True
          Case Is = 18   'top and right
              bchangeRight = True
              bchangeLeft = True
          Case Is = 20   'left and right
              bchangeRight = True
              bchangeTop = True
          Case Is = 22   'left and top and right
              bchangeRight = True
          Case Is = 24   'bottom and right
              bchangeRight = True
              bchangeBottom = True
              bchangeLeft = True
              bchangeTop = True
          Case Is = 26   'top and right and bottom
              bchangeRight = True
              bchangeBottom = True
              bchangeLeft = True
          Case Is = 28   'right and left and bottom
              bchangeRight = True
              bchangeTop = True
          Case Is = 30   'left and top and bottom and right
              bchangeRight = True
              bchangeBottom = True
        End Select
        
        'if one of the anchorpoints ISNT apLeft
        If bchangeLeft Then
           ichangeval = (.Parent.Width - iWid) + iLeft
           .Left = (ichangeval * imultiplyerLeft)
        End If
        
        'if one of the anchorpoints IS apRight
        If bchangeRight Then
           ichangeval = ((.Parent.Width - .Left) - iRight)
           If ichangeval < 0 Then
              ichangeval = 0
           End If
          .Width = ichangeval
        End If
        
        'if one of the anchorpoints ISNT top
        If bchangeTop Then
           ichangeval = ((.Parent.Height - iHei) + iTop)
          .Top = (ichangeval * imultiplyerTop)
        End If
        
        'if one of the anchorpoints IS apBottom
        If bchangeBottom Then
           ichangeval = ((.Parent.Height - .Top) - iBottom)
           If ichangeval < 0 Then
              ichangeval = 0
           End If
          .Height = ichangeval
        End If
        
        Set ctl = Nothing
      End With
    Next i
 
End Sub
 
 

Private Sub Class_Terminate()
   On Error Resume Next
   'stop subclassing
   SubClass.UnSubClass mForm.hwnd
   'erase array from memory
   Erase mCtrl
   Erase mCtrlVals
   Set mForm = Nothing
End Sub

 
