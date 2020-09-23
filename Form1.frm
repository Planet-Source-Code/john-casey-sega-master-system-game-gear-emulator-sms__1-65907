VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sega Master System Emulator"
   ClientHeight    =   2880
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   3840
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Machine 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   3240
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38& Then
                                controller1 = controller1 And Not 1&    '// Up
    ElseIf (KeyCode = 40&) Then controller1 = controller1 And Not 2&    '// Down
    ElseIf (KeyCode = 37&) Then controller1 = controller1 And Not 4&    '// Left
    ElseIf (KeyCode = 39&) Then controller1 = controller1 And Not 8&    '// Right
    ElseIf (KeyCode = 90&) Then controller1 = controller1 And Not 16&   '// Fire 1
    ElseIf (KeyCode = 88&) Then controller1 = controller1 And Not 32&   '// Fire 2
    End If
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 38& Then
                                controller1 = controller1 Or 1&         '// Up
    ElseIf (KeyCode = 40&) Then controller1 = controller1 Or 2&         '// Down
    ElseIf (KeyCode = 37&) Then controller1 = controller1 Or 4&         '// Left
    ElseIf (KeyCode = 39&) Then controller1 = controller1 Or 8&         '// Right
    ElseIf (KeyCode = 90&) Then controller1 = controller1 Or 16&        '// Fire 1
    ElseIf (KeyCode = 88&) Then controller1 = controller1 Or 32&        '// Fire 2
    End If

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    readCart (Data.Files.Item(1))
    Vdp.reset
    Z80.Z80Reset
   'Z80.execute
    Machine.Enabled = True      '// Calling Z80.execute from OLEDragDrop causes
                                '// Windows Explorer to 'hang' or 'freeze'.
                                '// This is a work around.
End Sub

Private Sub Machine_Timer()
    '// This is shit coding..
    Z80.execute: Machine.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
