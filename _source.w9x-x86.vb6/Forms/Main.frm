VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rope Demo"
   ClientHeight    =   4320
   ClientLeft      =   1350
   ClientTop       =   2295
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   502
   Begin VB.Frame fraDemoControls 
      Caption         =   "Demo Controls"
      Height          =   4275
      Left            =   4920
      TabIndex        =   0
      Top             =   30
      Width           =   2595
      Begin VB.CheckBox chkPointLocked 
         Caption         =   "Is Locked"
         Height          =   225
         Left            =   780
         TabIndex        =   7
         Top             =   540
         Width           =   1005
      End
      Begin VB.CommandButton cmdPoint 
         Caption         =   ">"
         Height          =   225
         Index           =   1
         Left            =   2310
         TabIndex        =   6
         Top             =   240
         Width           =   225
      End
      Begin VB.CommandButton cmdPoint 
         Caption         =   "<"
         Height          =   225
         Index           =   0
         Left            =   2070
         TabIndex        =   5
         Top             =   240
         Width           =   225
      End
      Begin VB.TextBox txtGravity 
         Height          =   315
         Left            =   870
         TabIndex        =   2
         Text            =   "9.81"
         Top             =   3870
         Width           =   885
      End
      Begin VB.Label lblPoint 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   4
         Top             =   270
         Width           =   90
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "Point:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   270
         Width           =   405
      End
      Begin VB.Label lblGravity 
         AutoSize        =   -1  'True
         Caption         =   "Gravity:"
         Height          =   195
         Left            =   270
         TabIndex        =   1
         Top             =   3930
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--VARIABLES--
' Private:
    Private i_gtDelta As Single

Public Sub Controls_SelectPoint(ByVal PointIndex As Long)
    On Error Resume Next
    
    lblPoint(1).Caption = PointIndex
    
    chkPointLocked.Value = IIf(RopeString.GetPointIsLocked(PointIndex - 1) = False, vbUnchecked, vbChecked)
End Sub
Private Function Controls_GetCurPoint() As Long
    On Error Resume Next
    
    Controls_GetCurPoint = CLng(lblPoint(1).Caption)
End Function

Private Sub chkPointLocked_Click()
    On Error Resume Next
    
    Dim dwPoint As Long
    dwPoint = Controls_GetCurPoint()
    If dwPoint < 1 Or dwPoint > RopeString.GetPointsCount() Then Exit Sub
    
    Call RopeString.SetPointIsLocked(dwPoint - 1, IIf(chkPointLocked.Value = vbUnchecked, False, True))
End Sub
Private Sub cmdPoint_Click(Index As Integer)
    On Error Resume Next
    Dim dwNext As Long
    
    Select Case Index
    Case 0  'Previous
        dwNext = (CLng(lblPoint(1).Caption) - 1)
        If dwNext < 1 Then dwNext = 1
        
        Call Controls_SelectPoint(dwNext)
    Case 1  'Next
        dwNext = (CLng(lblPoint(1).Caption) + 1)
        If dwNext > RopeString.GetPointsCount() Then dwNext = RopeString.GetPointsCount()
        
        Call Controls_SelectPoint(dwNext)
    End Select
End Sub


Private Sub Form_Load()
    On Error Resume Next
    
    ' Center this window.
    Me.Left = (Screen.Width - Me.Width) * 0.5
    Me.Top = (Screen.Height - Me.Height) * 0.5
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = vbRightButton Then
        Call RopeString.SetPointIsLocked(RopeString.GetPointsCount() - 1, True)
    End If
    
    Call Form_MouseMove(Button, Shift, x, y)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim fMouseX As Single, fMouseY As Single
    fMouseX = MAX_float(MIN_float(x, Me.fraDemoControls.Left), 0)
    fMouseY = MAX_float(MIN_float(y, Me.ScaleHeight), 0)
    
    If Button = vbLeftButton Then
        Call RopeString.SetPointPosition(0, fMouseX, (Me.ScaleHeight - fMouseY))
    ElseIf Button = vbRightButton Then
        Call RopeString.SetPointPosition(RopeString.GetPointsCount() - 1, fMouseX, (Me.ScaleHeight - fMouseY))
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = vbRightButton Then
        Call RopeString.SetPointIsLocked(RopeString.GetPointsCount() - 1, False)
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Call Loop_Terminate
    End If
End Sub

Private Sub txtGravity_Change()
    On Error Resume Next
    
    If Not txtGravity.Text = Val(txtGravity.Text) Then
        txtGravity.Text = Val(txtGravity.Text)
    End If
    
    Call RopeString.SetGravity(txtGravity.Text)
End Sub
