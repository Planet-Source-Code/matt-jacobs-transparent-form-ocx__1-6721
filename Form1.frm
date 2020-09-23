VERSION 5.00
Object = "{B5498E18-9F46-11D3-A30A-000001165224}#19.0#0"; "TransparentForm.ocx"
Begin VB.Form Form1 
   Caption         =   "Demo"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   1200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Untransparent"
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Container Transparent"
      Height          =   735
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Whole Transparent"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin TransparentForm.TranForm TranForm1 
      Left            =   2760
      Top             =   1200
      _extentx        =   1296
      _extenty        =   1296
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
TranForm1.SetWholeTransparent
End Sub

Private Sub Command2_Click()
TranForm1.SetUnTransparent

End Sub

Private Sub Command3_Click()
    TranForm1.SetContainerTransparent
End Sub
