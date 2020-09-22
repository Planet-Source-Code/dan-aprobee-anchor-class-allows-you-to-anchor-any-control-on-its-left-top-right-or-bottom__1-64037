VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4605
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   4110
      Begin MSFlexGridLib.MSFlexGrid FlexGrid1 
         Height          =   2085
         Left            =   90
         TabIndex        =   5
         Top             =   2385
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   3678
         _Version        =   393216
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   3660
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   285
         Left            =   2925
         TabIndex        =   2
         Top             =   315
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   315
         Width           =   2670
      End
      Begin VB.Label lblDivider 
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Left            =   0
         TabIndex        =   4
         Top             =   2250
         Width           =   4110
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsAnchor As C_anchor

Private Sub Form_Load()
Dim minwid  As Long
Dim maxwid  As Long
Dim minhei  As Long
Dim maxhei  As Long

   Set clsAnchor = New C_anchor
   
   'set forms min and max width and height
   minwid = (Me.Width - 1000)
   maxwid = (Me.Width + 2000)
   minhei = (Me.Height - 1000)
   maxhei = (Me.Height + 2000)
   
   With clsAnchor
     .SetFormDimensions Form1, minwid, maxwid, minhei, maxhei
     .Anchor Frame1, (apLeft Or apRight Or apTop Or apBottom)
     .Anchor Text1, (apRight Or apLeft Or apTop)
     .Anchor Command1, (apRight Or apTop)
     .Anchor List1, apAll
     .Anchor lblDivider, (apLeft Or apRight)
     .Anchor FlexGrid1, (apRight Or apLeft Or apBottom)
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set clsAnchor = Nothing
End Sub

Private Sub Form_Resize()
   clsAnchor.Parent_Resize_Event
End Sub

