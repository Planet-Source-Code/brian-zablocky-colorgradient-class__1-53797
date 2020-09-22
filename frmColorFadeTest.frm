VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmColorFadeTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ColorGradient Test"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmColorFadeTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frame 
      Caption         =   "Other Test"
      Height          =   1185
      Index           =   2
      Left            =   4725
      TabIndex        =   29
      Top             =   90
      Width           =   2130
      Begin VB.CommandButton cmdForm 
         Caption         =   "New Form"
         Height          =   330
         Left            =   450
         TabIndex        =   31
         Top             =   495
         Width           =   1230
      End
   End
   Begin VB.Frame Source 
      Caption         =   "Source"
      Height          =   1185
      Left            =   3420
      TabIndex        =   20
      Top             =   90
      Width           =   1185
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   8
         Left            =   495
         TabIndex        =   30
         ToolTipText     =   "Radial (Not Implemented)"
         Top             =   540
         Width           =   240
      End
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   7
         Left            =   225
         TabIndex        =   28
         ToolTipText     =   "CenterLeft"
         Top             =   540
         Value           =   -1  'True
         Width           =   240
      End
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   6
         Left            =   225
         TabIndex        =   27
         ToolTipText     =   "BottomLeft"
         Top             =   810
         Width           =   240
      End
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   5
         Left            =   495
         TabIndex        =   26
         ToolTipText     =   "BottomCenter"
         Top             =   810
         Width           =   240
      End
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   4
         Left            =   765
         TabIndex        =   25
         ToolTipText     =   "BottomRight"
         Top             =   810
         Width           =   240
      End
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   3
         Left            =   765
         TabIndex        =   24
         ToolTipText     =   "CenterRight"
         Top             =   540
         Width           =   240
      End
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   2
         Left            =   765
         TabIndex        =   23
         ToolTipText     =   "TopRight"
         Top             =   270
         Width           =   240
      End
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   1
         Left            =   495
         TabIndex        =   22
         ToolTipText     =   "TopCenter"
         Top             =   270
         Width           =   240
      End
      Begin VB.OptionButton optSource 
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   21
         ToolTipText     =   "TopLeft"
         Top             =   270
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   270
      Top             =   3330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frame 
      Caption         =   "Colors"
      Height          =   1185
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   3210
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   10
         Left            =   2790
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   9
         Left            =   2520
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   8
         Left            =   2250
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   11
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   7
         Left            =   1980
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   6
         Left            =   1710
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   5
         Left            =   1440
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   4
         Left            =   1170
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   3
         Left            =   900
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   6
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   630
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   5
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   360
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   90
         ScaleHeight     =   480
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   270
         Width           =   285
      End
      Begin VB.Label lblJunk 
         Alignment       =   2  'Center
         Caption         =   "80%"
         Height          =   240
         Index           =   5
         Left            =   2205
         TabIndex        =   19
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblJunk 
         Alignment       =   2  'Center
         Caption         =   "60%"
         Height          =   240
         Index           =   4
         Left            =   1665
         TabIndex        =   18
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblJunk 
         Alignment       =   2  'Center
         Caption         =   "40%"
         Height          =   240
         Index           =   3
         Left            =   1125
         TabIndex        =   17
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblJunk 
         Alignment       =   2  'Center
         Caption         =   "20%"
         Height          =   240
         Index           =   2
         Left            =   585
         TabIndex        =   16
         Top             =   810
         Width           =   375
      End
      Begin VB.Label lblJunk 
         Alignment       =   2  'Center
         Caption         =   "E"
         Height          =   240
         Index           =   1
         Left            =   2790
         TabIndex        =   14
         Top             =   810
         Width           =   285
      End
      Begin VB.Label lblJunk 
         Alignment       =   2  'Center
         Caption         =   "S"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   810
         Width           =   285
      End
   End
   Begin VB.Frame frame 
      Caption         =   "Sample"
      Height          =   2040
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   1350
      Width           =   6810
      Begin VB.PictureBox picTest 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400040&
         Height          =   1725
         Left            =   90
         ScaleHeight     =   1665
         ScaleWidth      =   6570
         TabIndex        =   1
         Top             =   225
         Width           =   6630
      End
   End
End
Attribute VB_Name = "frmColorFadeTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdForm_Click()
    frmTest.Show 1
End Sub

Private Sub Form_Load()
    
    CurrentSource = 7 'Center Left
    
    UpdatePictureBoxes
    
End Sub





Private Sub optSource_Click(Index As Integer)
    picTest.Cls
    Select Case Index
        Case 0: Gradient.PaintObject picTest, gTopLeft
        Case 1: Gradient.PaintObject picTest, gTopCenter
        Case 2: Gradient.PaintObject picTest, gTopRight
        Case 3: Gradient.PaintObject picTest, gCenterRight
        Case 4: Gradient.PaintObject picTest, gLowerRight
        Case 5: Gradient.PaintObject picTest, gLowerCenter
        Case 6: Gradient.PaintObject picTest, gLowerLeft
        Case 7: Gradient.PaintObject picTest, gCenterLeft
        Case 8: Gradient.PaintObject picTest, gCenterCenter
    End Select
    
    CurrentSource = Index
End Sub





Private Sub picColor_DblClick(Index As Integer)
    
    cdlg.Color = picColor(Index).BackColor
    
   'cdlg.DialogTitle = "Select " & Trim(Str(Index * 10)) & "% Color"
    cdlg.ShowColor
        
    If cdlg.Color <> picColor(Index).BackColor Then
        Gradient.SetColor Index, cdlg.Color
    End If
    
    UpdatePictureBoxes
    
End Sub






Private Sub UpdatePictureBoxes()
    Dim i As Integer
    For i = 0 To 10
        picColor(i).BackColor = Gradient.GetColor(i)
    Next i
        
    optSource_Click CurrentSource
    
End Sub
