VERSION 5.00
Begin VB.UserControl TranForm 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   735
   ScaleWidth      =   735
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "TranForm.ctx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "TranForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Const RGN_AND = 1
Const RGN_COPY = 5
Const RGN_DIFF = 4
Const RGN_OR = 2
Const RGN_XOR = 3


Type POINTAPI
    X As Long
    Y As Long
End Type


Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim DoAtStartup As Boolean


Public Sub SetWholeTransparent()
    Dim Frm As Form
    Dim rctClient As RECT, rctFrame As RECT
    Dim hClient As Long, hFrame As Long, hObj As Long
    Dim Start As Integer, Finish As Integer, I As Integer
    Dim InvisibleControl As Boolean
    
    Set Frm = UserControl.Parent
    
    GetWindowRect Frm.hwnd, rctFrame
    GetClientRect Frm.hwnd, rctClient
    
    Dim lpTL As POINTAPI, lpBR As POINTAPI
    lpTL.X = rctFrame.Left
    lpTL.Y = rctFrame.Top
    lpBR.X = rctFrame.Right
    lpBR.Y = rctFrame.Bottom
    ScreenToClient Frm.hwnd, lpTL
    ScreenToClient Frm.hwnd, lpBR
    rctFrame.Left = lpTL.X
    rctFrame.Top = lpTL.Y
    rctFrame.Right = lpBR.X
    rctFrame.Bottom = lpBR.Y
        
    rctFrame.Right = rctFrame.Right + Abs(rctFrame.Left)
    rctFrame.Bottom = rctFrame.Bottom + Abs(rctFrame.Top)
    rctFrame.Top = 0
    rctFrame.Left = 0
    
    hFrame = CreateRectRgn(rctFrame.Left, rctFrame.Top, rctFrame.Right, rctFrame.Bottom)
    Dim mode As Integer
    mode = Frm.ScaleMode
    Frm.ScaleMode = 3
    CombineRgn hFrame, hFrame, hFrame, RGN_XOR
    Start = 0
    Finish = Frm.Controls.Count - 1
    
    On Error GoTo Err_Hand
    For I = Start To Finish
        InvisibleControl = False
        hObj = CreateRectRgn(Frm.Controls(I).Left + 4, Frm.Controls(I).Top + 23, Frm.Controls(I).Left + Frm.Controls(I).Width + 4, Frm.Controls(I).Top + Frm.Controls(I).Height + 23)
        If InvisibleControl = False Then CombineRgn hFrame, hObj, hFrame, RGN_OR
    Next

    
    SetWindowRgn Frm.hwnd, hFrame, True
    Frm.ScaleMode = mode
    Exit Sub
Err_Hand:
    If Err.Number = 393 Then
        InvisibleControl = True
        Resume Next
    End If
End Sub


Public Sub SetContainerTransparent()
    Dim Frm As Form
    Dim rctClient As RECT, rctFrame As RECT
    Dim hClient As Long, hFrame As Long, hObj As Long
    Dim Start As Integer, Finish As Integer, I As Integer
    Dim InvisibleControl As Boolean
    
    Set Frm = UserControl.Parent

    GetWindowRect Frm.hwnd, rctFrame
    GetClientRect Frm.hwnd, rctClient
    
    Dim lpTL As POINTAPI, lpBR As POINTAPI
    lpTL.X = rctFrame.Left
    lpTL.Y = rctFrame.Top
    lpBR.X = rctFrame.Right
    lpBR.Y = rctFrame.Bottom
    ScreenToClient Frm.hwnd, lpTL
    ScreenToClient Frm.hwnd, lpBR
    rctFrame.Left = lpTL.X
    rctFrame.Top = lpTL.Y
    rctFrame.Right = lpBR.X
    rctFrame.Bottom = lpBR.Y
    rctClient.Left = Abs(rctFrame.Left)
    rctClient.Top = Abs(rctFrame.Top)
    rctClient.Right = rctClient.Right + Abs(rctFrame.Left)
    rctClient.Bottom = rctClient.Bottom + Abs(rctFrame.Top)
    rctFrame.Right = rctFrame.Right + Abs(rctFrame.Left) - 4
    rctFrame.Bottom = rctFrame.Bottom + Abs(rctFrame.Top) - 4
    rctFrame.Top = 4
    rctFrame.Left = 4
    
    hClient = CreateRectRgn(rctClient.Left, rctClient.Top, rctClient.Right, rctClient.Bottom)
    hFrame = CreateRectRgn(rctFrame.Left, rctFrame.Top, rctFrame.Right, rctFrame.Bottom)
    
    Dim mode As Integer
    mode = Frm.ScaleMode
    Frm.ScaleMode = 3
    CombineRgn hFrame, hClient, hFrame, RGN_XOR
    Start = 0
    Finish = Frm.Controls.Count - 1


    On Error GoTo Err_Hand
    For I = Start To Finish
        InvisibleControl = False
        hObj = CreateRectRgn(Frm.Controls(I).Left + 4, Frm.Controls(I).Top + 23, Frm.Controls(I).Left + Frm.Controls(I).Width + 4, Frm.Controls(I).Top + Frm.Controls(I).Height + 23)
        If InvisibleControl = False Then CombineRgn hFrame, hObj, hFrame, RGN_OR
    Next


    SetWindowRgn Frm.hwnd, hFrame, True
    Frm.ScaleMode = mode
    
    Exit Sub
Err_Hand:
    If Err.Number = 393 Then
        InvisibleControl = True
        Resume Next
    End If
    
End Sub

Public Sub SetUnTransparent()
    Dim Frm As Form
    
    Set Frm = UserControl.Parent
    
    SetWindowRgn Frm.hwnd, 0, True
End Sub


Private Sub UserControl_Resize()
    UserControl.Width = 735
    UserControl.Height = 735
    
End Sub


