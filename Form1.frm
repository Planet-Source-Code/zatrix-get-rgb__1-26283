VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   645
   End
   Begin VB.Label Label2 
      Height          =   930
      Left            =   120
      TabIndex        =   1
      Top             =   690
      Width           =   4545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Move the mouse anywhere on the screen to view the RGB color value of that pixel"
      Height          =   405
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Add a Timer Control To a Form, Then use this code And point To anywhere On your screen To have
'the RGB value appear In the Forms Caption.
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Form_Load()
    Timer1.Interval = 100
End Sub

Private Sub Timer1_Timer()
    Dim tPOS As POINTAPI
    Dim sTmp As String
    Dim lColor As Long
    Dim lDC As Long

    lDC = GetWindowDC(0)
    Call GetCursorPos(tPOS)
    lColor = GetPixel(lDC, tPOS.x, tPOS.y)
    Label2.BackColor = lColor
    
    sTmp = Right$("000000" & Hex(lColor), 6)
    Caption = "R:" & Right$(sTmp, 2) & " G:" & Mid$(sTmp, 3, 2) & " B:" & Left$(sTmp, 2)
End Sub
