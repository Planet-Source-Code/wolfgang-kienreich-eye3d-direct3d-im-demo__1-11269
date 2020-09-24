VERSION 5.00
Begin VB.Form fMenu 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "fMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Eye3D"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "fMenu.frx":0442
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   Picture         =   "fMenu.frx":074C
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   16
      X1              =   237
      X2              =   251
      Y1              =   384
      Y2              =   384
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   420
      Index           =   15
      Left            =   2265
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   1290
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   330
      Index           =   15
      Left            =   2385
      TabIndex        =   25
      Top             =   5445
      Width           =   1050
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   14
      X1              =   365
      X2              =   374
      Y1              =   384
      Y2              =   384
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   7
      X1              =   256
      X2              =   277
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   14
      Left            =   4155
      Shape           =   4  'Rounded Rectangle
      Top             =   1050
      Width           =   3165
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Not detected"
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
      Height          =   225
      Index           =   14
      Left            =   4275
      TabIndex        =   24
      Top             =   1110
      Width           =   2925
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H00404040&
      Index           =   3
      X1              =   499
      X2              =   499
      Y1              =   402
      Y2              =   2
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H00404040&
      Index           =   2
      X1              =   0
      X2              =   499
      Y1              =   399
      Y2              =   399
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   399
   End
   Begin VB.Line lneBorder 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   1
      X2              =   500
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENABLED"
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
      Index           =   13
      Left            =   5220
      TabIndex        =   22
      Top             =   3735
      Width           =   2025
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   13
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   3690
      Width           =   2175
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Specular lighting"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   8
      Left            =   2235
      TabIndex        =   23
      Top             =   3780
      Width           =   1995
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   15
      X1              =   148
      X2              =   345
      Y1              =   266
      Y2              =   266
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   13
      X1              =   10
      X2              =   151
      Y1              =   384
      Y2              =   384
   End
   Begin VB.Label lblStats 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxx"
      ForeColor       =   &H00004080&
      Height          =   210
      Left            =   150
      TabIndex        =   21
      Top             =   5550
      Width           =   1800
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "||||||||||||||||||||"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   10
      Left            =   4980
      TabIndex        =   18
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   330
      Index           =   12
      Left            =   5700
      TabIndex        =   20
      Top             =   5445
      Width           =   1575
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   330
      Index           =   11
      Left            =   3885
      TabIndex        =   19
      Top             =   5445
      Width           =   1485
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   12
      X1              =   416
      X2              =   433
      Y1              =   331
      Y2              =   331
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   11
      X1              =   313
      X2              =   329
      Y1              =   331
      Y2              =   331
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   10
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   4650
      Width           =   1335
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Caption         =   "SMALL"
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
      Index           =   9
      Left            =   3885
      TabIndex        =   17
      Top             =   4695
      Width           =   690
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   10
      X1              =   150
      X2              =   256
      Y1              =   331
      Y2              =   331
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LARGE"
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
      Index           =   8
      Left            =   6570
      TabIndex        =   16
      Top             =   4695
      Width           =   690
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen size"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   6
      Left            =   2250
      TabIndex        =   15
      Top             =   4740
      Width           =   1275
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   5
      X1              =   149
      X2              =   346
      Y1              =   295
      Y2              =   295
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENABLED"
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
      Index           =   5
      Left            =   5235
      TabIndex        =   14
      Top             =   4170
      Width           =   2025
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Phong shading"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   7
      Left            =   2235
      TabIndex        =   13
      Top             =   4215
      Width           =   1995
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   9
      X1              =   256
      X2              =   256
      Y1              =   88
      Y2              =   32
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   8
      X1              =   256
      X2              =   277
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "HAL optimized hardware device"
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
      Height          =   225
      Index           =   7
      Left            =   4275
      TabIndex        =   12
      Top             =   690
      Width           =   2925
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RGB software emulation device"
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
      Height          =   225
      Index           =   6
      Left            =   4275
      TabIndex        =   11
      Top             =   255
      Width           =   2925
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   6
      X1              =   147
      X2              =   277
      Y1              =   33
      Y2              =   33
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   4
      X1              =   148
      X2              =   345
      Y1              =   238
      Y2              =   238
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENABLED"
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
      Index           =   4
      Left            =   5220
      TabIndex        =   9
      Top             =   3315
      Width           =   2025
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   3
      X1              =   148
      X2              =   345
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENABLED"
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
      Index           =   3
      Left            =   5220
      TabIndex        =   8
      Top             =   2895
      Width           =   2025
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   2
      X1              =   148
      X2              =   345
      Y1              =   182
      Y2              =   182
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENABLED"
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
      Index           =   2
      Left            =   5220
      TabIndex        =   7
      Top             =   2475
      Width           =   2025
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   1
      X1              =   148
      X2              =   345
      Y1              =   154
      Y2              =   154
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENABLED"
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
      Index           =   1
      Left            =   5220
      TabIndex        =   6
      Top             =   2055
      Width           =   2025
   End
   Begin VB.Line lneCaption 
      BorderColor     =   &H00004080&
      Index           =   0
      X1              =   148
      X2              =   345
      Y1              =   126
      Y2              =   126
   End
   Begin VB.Label lblButton 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ENABLED"
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
      Index           =   0
      Left            =   5220
      TabIndex        =   5
      Top             =   1635
      Width           =   2025
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Linear texture filtering"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   4
      Left            =   2235
      TabIndex        =   4
      Top             =   3360
      Width           =   1995
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Texture translucency"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   3
      Left            =   2235
      TabIndex        =   3
      Top             =   2940
      Width           =   1995
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Texture transparency"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   2
      Left            =   2235
      TabIndex        =   2
      Top             =   2520
      Width           =   1995
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Perspectivic correction"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   1
      Left            =   2235
      TabIndex        =   1
      Top             =   2100
      Width           =   1995
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Texture mapping"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   0
      Left            =   2235
      TabIndex        =   0
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Direct3D driver"
      ForeColor       =   &H00004080&
      Height          =   270
      Index           =   5
      Left            =   2235
      TabIndex        =   10
      Top             =   285
      Width           =   1500
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   420
      Index           =   12
      Left            =   5610
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   1725
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   420
      Index           =   11
      Left            =   3765
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   1725
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   9
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   4650
      Width           =   855
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   8
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   4650
      Width           =   855
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   5
      Left            =   5175
      Shape           =   4  'Rounded Rectangle
      Top             =   4125
      Width           =   2160
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   7
      Left            =   4155
      Shape           =   4  'Rounded Rectangle
      Top             =   630
      Width           =   3165
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   4
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   3270
      Width           =   2160
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   3
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   2850
      Width           =   2160
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   2
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   2430
      Width           =   2160
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   1
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   2010
      Width           =   2160
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   0
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   1590
      Width           =   2160
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      Height          =   345
      Index           =   6
      Left            =   4155
      Shape           =   4  'Rounded Rectangle
      Top             =   210
      Width           =   3165
   End
End
Attribute VB_Name = "fMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Color constants
Private Const RGBHighLight = &H80FF&       ' Highlight for mouse hover
Private Const RGBStandard = &H4080&        ' Default color
Private Const RGBSelected = &HC0C0&        ' Highlight for selected items
Private Const RGBBack = &H0&               ' Background color


' FORM_LOAD: Reset all controls to default color
Private Sub Form_Load()
        
    ' Initialize controls using hardware found
    Call InitControls
    
    ' Update controls for seleced driver
    Call UpdateControls
    
    ' Initialize control drawstate
    Call RedrawHighlights(-1)
    
End Sub

' FORM_MOUSEMOVE: Highlight handling for empty form areas
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call RedrawHighlights(-1)
End Sub
' LBLBUTTON_MOUSEDOWN: Click handling for button label controls
Private Sub lblButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Extremely highlight selected control to display click
    Me.lblButton(Index).ForeColor = RGBBack
    Me.shpButton(Index).FillColor = RGBStandard
    Me.shpButton(Index).FillStyle = 0
    
    ' React to click
    Select Case Index
         
        ' Texturemapping ?
        Case 0
            If Me.lblButton(0).Caption <> "Not supported" Then
                If Me.lblButton(0).Caption = "ENABLED" Then
                    Me.lblButton(0).Caption = "DISABLED"
                    G_dUser.DisplayOptions.Mapping = False
                Else
                    Me.lblButton(0).Caption = "ENABLED"
                    G_dUser.DisplayOptions.Mapping = True
                End If
            End If
            
        ' Texture perspectivic correction ?
        Case 1
            If Me.lblButton(1).Caption <> "Not supported" Then
                If Me.lblButton(1).Caption = "ENABLED" Then
                    Me.lblButton(1).Caption = "DISABLED"
                    G_dUser.DisplayOptions.Correct = False
                Else
                    Me.lblButton(1).Caption = "ENABLED"
                    G_dUser.DisplayOptions.Correct = True
                End If
            End If
        
        ' Transparent (color-keyed) textures ?
        Case 2
            If Me.lblButton(2).Caption <> "Not supported" Then
                If Me.lblButton(2).Caption = "ENABLED" Then
                    Me.lblButton(2).Caption = "DISABLED"
                    G_dUser.DisplayOptions.Transparent = False
                Else
                    Me.lblButton(2).Caption = "ENABLED"
                    G_dUser.DisplayOptions.Transparent = True
                End If
            End If
        
        ' Translucent (alpha-blended textures ?
        Case 3
            If Me.lblButton(3).Caption <> "Not supported" Then
                If Me.lblButton(3).Caption = "ENABLED" Then
                    Me.lblButton(3).Caption = "DISABLED"
                    G_dUser.DisplayOptions.Translucent = False
                Else
                    Me.lblButton(3).Caption = "ENABLED"
                    G_dUser.DisplayOptions.Translucent = True
                End If
            End If
        
        ' Bilinear filtering ?
        Case 4
            If Me.lblButton(4).Caption <> "Not supported" Then
                If Me.lblButton(4).Caption = "ENABLED" Then
                    Me.lblButton(4).Caption = "DISABLED"
                    G_dUser.DisplayOptions.Filtering = False
                Else
                    Me.lblButton(4).Caption = "ENABLED"
                    G_dUser.DisplayOptions.Filtering = True
                End If
            End If
        
        ' Phong ?
        Case 5
            If Me.lblButton(5).Caption <> "Not supported" Then
                If Me.lblButton(5).Caption = "ENABLED" Then
                    Me.lblButton(5).Caption = "DISABLED"
                    G_dUser.DisplayOptions.Phong = False
                Else
                    Me.lblButton(5).Caption = "ENABLED"
                    G_dUser.DisplayOptions.Phong = True
                End If
            End If
        
        ' Specular highlights ?
        Case 13
            If Me.lblButton(13).Caption <> "Not supported" Then
                If Me.lblButton(13).Caption = "ENABLED" Then
                    Me.lblButton(13).Caption = "DISABLED"
                    G_dUser.DisplayOptions.Specular = False
                Else
                    Me.lblButton(13).Caption = "ENABLED"
                    G_dUser.DisplayOptions.Specular = True
                End If
            End If
        
        
        ' Select RGB driver
        Case 6
            If G_dDXDriverSoft.Found Then
                G_dDXSelectedDriver = G_dDXDriverSoft
                UpdateControls
            End If
            
        ' Select HAL driver
        Case 7
            If G_dDXDriverHard.Found Then
                G_dDXSelectedDriver = G_dDXDriverHard
                UpdateControls
            End If
            
        ' Select 3DFX driver
        Case 14
            If G_dDXDriverPlus.Found Then
                G_dDXSelectedDriver = G_dDXDriverPlus
                UpdateControls
            End If
        
        ' Larger
        Case 8
            If Len(Me.lblButton(10).Caption) < 20 Then
                Me.lblButton(10).Caption = Me.lblButton(10).Caption & "|"
                G_dUser.DisplaySize = (20 - Len(Me.lblButton(10).Caption)) * 10
            End If
            
        ' Smaller
        Case 9
            If Len(Me.lblButton(11).Caption) > 1 Then
                Me.lblButton(10).Caption = Left(Me.lblButton(10).Caption, Len(Me.lblButton(10).Caption) - 1)
                G_dUser.DisplaySize = (20 - Len(Me.lblButton(10).Caption)) * 10
            End If
        
        ' Quit
        Case 11
            Me.Hide
            On Error Resume Next
            Shell App.Path + "\nls.exe", vbNormalFocus
            Unload Me
            End
            
        ' Start
        Case 12
            Unload Me
            
        ' Info
        Case 15
            
            Me.Hide
            fMsg.Hide
            fMsg.lblTitle = "EYE3D (C) 1999 by Nonlinear Solutions"
            fMsg.lblText = "Use arrow keys to move, mouse to look, ESC to exit." + vbCrLf + vbCrLf + "Known bugs: On 3DFX cards, texture animations don't work. On some 2D cards, translucency will be listed available, but translucent surfaces will not be visible." + vbCrLf + vbCrLf + "Visit us at www.dige.com/nls"
            fMsg.Show 1
            Me.Show 1
        
    End Select
    
End Sub
' LBLBUTTON_MOUSEMOVE: Highlight handling for button label controls
Private Sub lblButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then Call RedrawHighlights(Index)
End Sub
' LBLBUTTON_MOUSEUP: Highlight handling for button label control
Private Sub lblButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RedrawHighlights(-1)
End Sub
' LBLCAPTION_MOUSEMOVE: Highlight handling for caption controls
Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RedrawHighlights(-1)
End Sub

' REDRAWHIGHLIGHTS: Redraws highlights for controls
Private Sub RedrawHighlights(nControlIndex As Integer)

    ' Setup local variables ...
        Dim nRunControls As Integer ' Variable to run through controls
        
    ' Redraw highlights ...
        
        For nRunControls = 0 To 15
            If nRunControls = 6 Then
                Me.lblButton(nRunControls).ForeColor = IIf(G_dDXSelectedDriver.DriverType = EDXDTSoft, RGBSelected, IIf(nControlIndex = 6, RGBHighLight, RGBStandard))
            ElseIf nRunControls = 7 Then
                Me.lblButton(nRunControls).ForeColor = IIf(G_dDXSelectedDriver.DriverType = EDXDTHard, RGBSelected, IIf(nControlIndex = 7, RGBHighLight, RGBStandard))
            ElseIf nRunControls = 14 Then
                Me.lblButton(nRunControls).ForeColor = IIf(G_dDXSelectedDriver.DriverType = EDXDTPlus, RGBSelected, IIf(nControlIndex = 14, RGBHighLight, RGBStandard))
            Else
                Me.lblButton(nRunControls).ForeColor = RGBStandard
            End If
            Me.shpButton(nRunControls).FillStyle = 1
            Me.shpButton(nRunControls).BorderColor = IIf(nRunControls = nControlIndex, RGBHighLight, RGBStandard)
        Next
        
        For nRunControls = 0 To Me.lneCaption.Count - 1
            Me.lneCaption(nRunControls).BorderColor = RGBStandard
        Next
        
        For nRunControls = 0 To Me.lblCaption.Count - 1
            Me.lblCaption(nRunControls).ForeColor = RGBStandard
        Next
        
End Sub

' UPDATECONTROLS: Refill control captions with values based on driver selection
Private Sub UpdateControls()

    
    ' THE FOLLOWING FEATURES ARE DRIVER-DEPENDEND AND ARE THEREFORE CHECKED AGAINST THE DRIVER DESCRIPTION ...
    
        ' Get settings from driver currently selected
        With G_dDXSelectedDriver.D3DDriver
            
            ' Bilinear filtering ?
            If (.DEVDESC.dpcTriCaps.dwTextureFilterCaps And D3DPTFILTERCAPS_LINEAR) = D3DPTFILTERCAPS_LINEAR Then
                If G_dUser.DisplayOptions.Filtering Then
                    Me.lblButton(4).Caption = "ENABLED"
                Else
                    Me.lblButton(4).Caption = "DISABLED"
                End If
            Else
                Me.lblButton(4).Caption = "Not supported"
                G_dUser.DisplayOptions.Filtering = False
            End If
            
            ' Transparent (color-keyed) textures ?
            If (.DEVDESC.dpcTriCaps.dwTextureCaps And D3DPTEXTURECAPS_TRANSPARENCY) = D3DPTEXTURECAPS_TRANSPARENCY Then
                If G_dUser.DisplayOptions.Transparent Then
                    Me.lblButton(2).Caption = "ENABLED"
                Else
                    Me.lblButton(2).Caption = "DISABLED"
                End If
            Else
                Me.lblButton(2).Caption = "Not supported"
                G_dUser.DisplayOptions.Transparent = False
            End If
            
            ' Translucent (alpha-blended) textures ?
            If ((.DEVDESC.dpcTriCaps.dwSrcBlendCaps And D3DPBLENDCAPS_BOTHSRCALPHA) = D3DPBLENDCAPS_BOTHSRCALPHA) Then
                If G_dUser.DisplayOptions.Translucent Then
                    Me.lblButton(3).Caption = "ENABLED"
                Else
                    Me.lblButton(3).Caption = "DISABLED"
                End If
            Else
                Me.lblButton(3).Caption = "Not supported"
                G_dUser.DisplayOptions.Translucent = False
            End If
            
            ' Phong shading
            If ((.DEVDESC.dpcTriCaps.dwShadeCaps And D3DPSHADECAPS_COLORPHONGRGB) = D3DPSHADECAPS_COLORPHONGRGB) Then
                If G_dUser.DisplayOptions.Phong Then
                    Me.lblButton(5).Caption = "ENABLED"
                Else
                    Me.lblButton(5).Caption = "DISABLED"
                End If
            Else
                Me.lblButton(5).Caption = "Not supported"
                G_dUser.DisplayOptions.Phong = False
            End If
            
            ' Specular lighting ?
            If (.DEVDESC.dpcTriCaps.dwShadeCaps And D3DPSHADECAPS_SPECULARGOURAUDRGB) = D3DPSHADECAPS_SPECULARGOURAUDRGB Then
                If G_dUser.DisplayOptions.Specular Then
                    Me.lblButton(13).Caption = "ENABLED"
                Else
                    Me.lblButton(13).Caption = "DISABLED"
                End If
            Else
                Me.lblButton(13).Caption = "Not supported"
                G_dUser.DisplayOptions.Specular = False
            End If
        
        End With
    
    ' THE FOLLOWING FEATURES SEEM TO WORK ON ANY CARD ...
    
        ' Texturemapping ?
        If G_dUser.DisplayOptions.Mapping Then
            Me.lblButton(0).Caption = "ENABLED"
        Else
            Me.lblButton(0).Caption = "DISABLED"
        End If
        
        ' Perspectivic correction ?
        If G_dUser.DisplayOptions.Correct Then
            Me.lblButton(1).Caption = "ENABLED"
        Else
            Me.lblButton(1).Caption = "DISABLED"
        End If
    
    ' MISCELLANEOUS FEATURES ...
    
        ' Screensize
        Me.lblButton(10).Caption = String(20 - G_dUser.DisplaySize \ 20, "|")
    
End Sub

' INITCONTROLS: Initialized controls based on found hardware, write stats
Private Sub InitControls()
    
    ' STATS ...

        Me.lblStats.Caption = "Avg. frametime: " & IIf(G_dUser.Stats.Frametime < 1, "n/a", Format(G_dUser.Stats.Frametime, "0.0") & "fps")
        
    ' HAL driver...
        If G_dDXDriverHard.Found Then
            Me.lblButton(7).Caption = "HAL optimized hardware device"
        Else
            Me.lblButton(7).Caption = "No HAL device driver detected"
        End If
        
    ' Hardware accellerator...
        If G_dDXDriverPlus.Found Then
            Me.lblButton(14).Caption = Left(G_dDXDriverPlus.Name, 30)
        Else
            Me.lblButton(14).Caption = "No add-on board detected"
        End If
        
End Sub

