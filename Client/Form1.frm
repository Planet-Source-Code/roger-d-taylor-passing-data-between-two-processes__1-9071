VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   285
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Left            =   180
      TabIndex        =   1
      Top             =   825
      Width           =   2835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    gHW = Me.hwnd
    Hook
    Me.Caption = "Target"
    Me.Show
    Label1.Caption = Hex$(gHW)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unhook
End Sub
