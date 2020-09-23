VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " XP Style Toolbar-Lite"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pcToolBar 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6960
      TabIndex        =   0
      Top             =   0
      Width           =   6960
      Begin VB.Shape shpShadowBottom 
         BorderColor     =   &H00C00000&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   330
         Left            =   210
         Top             =   75
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.Shape shpShadowRight 
         BorderColor     =   &H00C00000&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   330
         Left            =   420
         Top             =   30
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.Shape shpMover 
         BorderColor     =   &H00FF8080&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   5805
         Top             =   435
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   11
         Left            =   5265
         Picture         =   "Form1.frx":000C
         Tag             =   "ALIGNRIGHT"
         Top             =   645
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   10
         Left            =   4800
         Picture         =   "Form1.frx":0156
         Tag             =   "ALIGNCENTER"
         Top             =   630
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   9
         Left            =   4500
         Picture         =   "Form1.frx":02A0
         Tag             =   "ALIGNLEFT"
         Top             =   630
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   8
         Left            =   4065
         Picture         =   "Form1.frx":03EA
         Tag             =   "UNDERLINE"
         Top             =   645
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   7
         Left            =   3615
         Picture         =   "Form1.frx":0534
         Tag             =   "ITALIC"
         Top             =   735
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   6
         Left            =   3300
         Picture         =   "Form1.frx":067E
         Tag             =   "BOLD"
         Top             =   735
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   5
         Left            =   2190
         Picture         =   "Form1.frx":07C8
         Tag             =   "PASTE"
         Top             =   750
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   4
         Left            =   1815
         Picture         =   "Form1.frx":0912
         Tag             =   "COPY"
         Top             =   675
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   3
         Left            =   1080
         Picture         =   "Form1.frx":0A5C
         Tag             =   "CUT"
         Top             =   795
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   2
         Left            =   750
         Picture         =   "Form1.frx":0BA6
         Tag             =   "SAVE"
         Top             =   810
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   1
         Left            =   450
         Picture         =   "Form1.frx":0CF0
         Tag             =   "OPEN"
         Top             =   825
         Width           =   240
      End
      Begin VB.Image imgButton 
         Height          =   240
         Index           =   0
         Left            =   135
         Picture         =   "Form1.frx":127A
         Tag             =   "NEW"
         ToolTipText     =   "New"
         Top             =   795
         Width           =   240
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Please send comments to: ramandy@rediffmail.com"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1365
      Width           =   3705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "No hooks or API. Only Shape and Image controls used."
      Height          =   195
      Left            =   540
      TabIndex        =   2
      Top             =   1065
      Width           =   3960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Simulates Office XP Style Toolbar Buttons."
      Height          =   195
      Left            =   540
      TabIndex        =   1
      Top             =   750
      Width           =   3030
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
': Module:          XPStyle Toolbar
': Description:     Make XP style toolbar with vb intrincic controls
': Comments to:     ramandy@rediffmail.com
':                  <I'm waiting for ur comments>
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Option Explicit

': Default Button Panel Height
Private Const IntTopPanelHeight As Integer = 450

Private Sub Form_Load()

    shpShadowBottom.ZOrder
    shpShadowRight.ZOrder
    ': Set colors u like
    With shpMover
        .ZOrder
        .FillColor = &HFFC0C0
        .BorderColor = &HFF8080
        .DrawMode = 9   ': Mask Pen, !!-IMPORTANT-!!
    End With
    
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
    ': Move our panel
    pcToolBar.Move 0, 0, Me.ScaleWidth, IntTopPanelHeight
End Sub

Private Sub imgButton_Click(Index As Integer)
': We've provided tag property for each buttons
': to identify each buttons.
    
    shpShadowRight.Visible = False
    shpShadowBottom.Visible = False
    shpMover.Visible = False
    
    With imgButton(Index)
        Select Case Trim(UCase(imgButton(Index).Tag))
            Case "NEW"
                ': Code here
            Case "OPEN"
                ': Code here
            ': .............
        End Select

    MsgBox "You clicked " & .Tag, vbInformation, App.Title
    End With


End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpShadowRight.Visible = False
    shpShadowBottom.Visible = False
    shpMover.Move shpMover.Left + 15, shpMover.Top + 15
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgButton(Index)
        shpMover.Move .Left - 60, .Top - 60, .Height + 120, .Width + 120
        shpShadowBottom.Move shpMover.Left + 30, (shpMover.Top + shpMover.Height), shpMover.Width - 15, 15
        shpShadowRight.Move (shpMover.Left + shpMover.Width), shpMover.Top + 30, 15, shpMover.Height - 15
        shpShadowRight.Visible = True
        shpShadowBottom.Visible = True
        shpMover.Visible = True
    End With
End Sub

Private Sub pcToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If shpMover.Visible Then shpMover.Visible = False
    If shpShadowRight.Visible Then shpShadowRight.Visible = False
    If shpShadowBottom.Visible Then shpShadowBottom.Visible = False
End Sub

Private Sub pcToolBar_Resize()
On Local Error Resume Next
Dim IntI As Integer
Dim IntLeft As Integer

Const IntWidth As Integer = 390 ': Default Button Width+Seperator width
Const IntTop As Integer = 90    ': Default Button top


IntLeft = 150                   ': Initial Button left

': Position buttons
    For IntI = imgButton.LBound To imgButton.UBound
        imgButton(IntI).Move IntLeft, IntTop
        IntLeft = IntLeft + IntWidth
    Next IntI

End Sub
