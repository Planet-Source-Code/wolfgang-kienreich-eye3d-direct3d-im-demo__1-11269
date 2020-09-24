VERSION 5.00
Begin VB.Form fMsg 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "fMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "fMsg.frx":0442
   MousePointer    =   99  'Custom
   Picture         =   "fMsg.frx":074C
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   4020
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   825
      Width           =   4020
   End
   Begin VB.Shape shpOkay 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Left            =   2535
      Shape           =   4  'Rounded Rectangle
      Top             =   2505
      Width           =   2160
   End
   Begin VB.Label lblOkay 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OKAY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   270
      Left            =   2595
      TabIndex        =   0
      Top             =   2550
      Width           =   2025
   End
   Begin VB.Line lneOkay 
      BorderColor     =   &H00004080&
      X1              =   10
      X2              =   170
      Y1              =   187
      Y2              =   187
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H00404040&
      Index           =   3
      X1              =   0
      X2              =   319
      Y1              =   199
      Y2              =   199
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H00404040&
      Index           =   2
      X1              =   319
      X2              =   319
      Y1              =   0
      Y2              =   200
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   1
      X2              =   320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   200
   End
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Color constants
Private Const RGBHighLight = &H80FF&       ' Highlight for mouse hover
Private Const RGBStandard = &H4080&        ' Default color

' FORM_MOUSEMOVE: Highlight handling for empty form areas
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Me.shpOkay.BorderColor = RGBStandard
End Sub
' LBLOKAY_MOUSEDOWN: Click handling for button label controls
Private Sub lblOkay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
End Sub
' LBLOKAY_MOUSEMOVE: Highlight handling for button
Private Sub lblOkay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.shpOkay.BorderColor = RGBHighLight
End Sub
