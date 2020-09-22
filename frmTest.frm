VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Resizable ColorGradient Test"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    
     Select Case CurrentSource
        Case 0: Gradient.PaintObject Me, gTopLeft
        Case 1: Gradient.PaintObject Me, gTopCenter
        Case 2: Gradient.PaintObject Me, gTopRight
        Case 3: Gradient.PaintObject Me, gCenterRight
        Case 4: Gradient.PaintObject Me, gLowerRight
        Case 5: Gradient.PaintObject Me, gLowerCenter
        Case 6: Gradient.PaintObject Me, gLowerLeft
        Case 7: Gradient.PaintObject Me, gCenterLeft
        Case 8: Gradient.PaintObject Me, gCenterCenter
    End Select
        
    Me.Refresh
    
End Sub
