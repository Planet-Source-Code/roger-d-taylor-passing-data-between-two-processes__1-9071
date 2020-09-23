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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1740
      TabIndex        =   0
      Top             =   1335
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Const WM_COPYDATA = &H4A
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Copies a block of memory from one location to another.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Sub Command1_Click()
    Dim cdCopyData As COPYDATASTRUCT
        Dim ThWnd As Long
    Dim byteBuffer(1 To 255) As Byte
    Dim strTemp As String
    
    ' Get the hWnd of the target application
        ThWnd = FindWindow(vbNullString, "Target")
        strTemp = "This is the data that will be sent"
    
        ' Copy the string into a byte array, converting it to ASCII
    Call CopyMemory(byteBuffer(1), ByVal strTemp, Len(strTemp))
        cdCopyData.dwData = 3
        cdCopyData.cbData = Len(strTemp) + 1
        cdCopyData.lpData = VarPtr(byteBuffer(1))
        i = SendMessage(ThWnd, WM_COPYDATA, Me.hwnd, cdCopyData)
        
End Sub


Private Sub Form_Load()
      ' This gives you visibility that the target app is running
      ' and you are pointing to the correct hWnd
      Me.Caption = Hex$(FindWindow(vbNullString, "Target"))
End Sub


