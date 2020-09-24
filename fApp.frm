VERSION 5.00
Begin VB.Form fApp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
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
   ForeColor       =   &H00000000&
   Icon            =   "fApp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Eye3D"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "fApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' FORM_ACTIVATE: Initialize application, start main loop
Private Sub Form_Activate()

    ' Only do this once
    If Not G_bAppInitialized Then
        
        ' Set initialisation status
        G_bAppInitialized = True         ' Initialize and start
        
        ' Initialize user data
        With G_dUser
            .Position.X = 57
            .Position.Y = 43
            .Position.z = 44
            .LookH = 180
            .LookV = -12
            .Speed = 0.5
            .DisplaySize = 1
            .DisplayOptions.Correct = True
            .DisplayOptions.Filtering = True
            .DisplayOptions.Phong = False
            .DisplayOptions.Mapping = True
            .DisplayOptions.Specular = True
            .DisplayOptions.Translucent = True
            .DisplayOptions.Transparent = True
        End With
        
        ' Initialize display size
        G_nDisplayWidth = 640
        G_nDisplayHeight = 480
        
        ' Detect DirectX caps
        Call AppDriverDetect
        
        ' Do main loop (menu/app)
        Do
            
            ' Show menu
            Me.Hide
            fMenu.Show 1
            Me.Show
            
           ' Show demo form
            Call AppInitialize
            Call AppLoop

        Loop
        
    End If
    
End Sub

' FORM_KEYUP: Reset keystate
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    G_dUser.InputState.KeyCode = 0
End Sub

' MOUSE_MOVE: Turn "head" according to relative change in mouse position
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    G_dUser.InputState.MouseButton = Button
    G_dUser.InputState.MouseX = X
    G_dUser.InputState.MouseY = Y
    
End Sub

' FORM_UNLOAD: Emergency shutdown of DirectX objects
Private Sub Form_Unload(Cancel As Integer)

    ' Emergency shutdown
    AppTerminate
    End
    
End Sub

' FORM_KEYDOWN: React to user key input
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    G_dUser.InputState.KeyCode = KeyCode
End Sub

