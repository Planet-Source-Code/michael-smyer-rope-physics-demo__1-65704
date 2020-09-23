Attribute VB_Name = "basMain"
Option Explicit

'--VARIABLES--
' Private:
    Private i_End As Boolean
    Private i_LastTickTime As Long  'For calculate delta time.
' Public:
    Public e_gtDelta As Single
' Interfaces:
    Public RopeString As clsString

'--API DECLARATIONS--
' Private:
    Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Function App_CalcDeltaTime() As Single
    On Error Resume Next
    
    Dim dwCurTime As Long
    dwCurTime = GetTickCount()
    
    Dim dwTickDiff As Long
    dwTickDiff = (dwCurTime - i_LastTickTime)
    
    App_CalcDeltaTime = (CSng(dwTickDiff) / 1000#)
    If App_CalcDeltaTime > 0.05 Then App_CalcDeltaTime = 0.05
    
    i_LastTickTime = dwCurTime
End Function
Private Sub App_End()
    On Error Resume Next
    
    ' Terminate this application.
    End
End Sub
Private Sub Loop_Continue()
    On Error Resume Next
    
    ' Get the global delta time.
    e_gtDelta = App_CalcDeltaTime()
    
    ' Continue all events.
    Call Loop_DoEvents
    
    ' Display the contents.
    Call Loop_Display
End Sub
Private Sub Loop_DoEvents()
    On Error Resume Next
    
    ' Continue the rope.
    Call RopeString.Continue(e_gtDelta)
    
    ' Display the statistics.
    'Dim strNewCaption As String
    'strNewCaption = Round(RopeString.GetCurLength(), 0) & "px"
    'If Not frmMain.lblCurLength(1).Caption = strNewCaption Then frmMain.lblCurLength(1).Caption = strNewCaption
End Sub
Private Sub Loop_Start()
    On Error Resume Next
    
    ' Make sure the loop doesn't end at the start.
    i_End = False
    
    ' Start the loop ...
    Do
        ' ... If the application has focus then ...
        If frmMain.WindowState <> vbMinimized Then
            ' ... Continue all events ...
            Call Loop_Continue
        End If
        
        ' ... Continue all system events ...
        DoEvents
    ' ... Continue the loop until it is told to end.
    Loop Until i_End = True
    
    ' Terminate the application.
    Call App_End
End Sub
Private Sub Loop_Display()
    On Error Resume Next
    
    Call frmMain.Cls
    
    Dim dwIndex As Long, fX1 As Single, fY1 As Single, fX2 As Single, fY2 As Single
    For dwIndex = 0 To (RopeString.GetPointsCount() - 1) - 1  'don't include the last item because of how the line is drawn.
        Call RopeString.GetPointPosition(dwIndex, fX1, fY1)
        Call RopeString.GetPointPosition(dwIndex + 1, fX2, fY2)
        
        frmMain.Line (fX1, (frmMain.ScaleHeight - fY1))-(fX2, (frmMain.ScaleHeight - fY2)), vbRed
    Next dwIndex
    
    Call frmMain.Refresh
End Sub
Public Sub Loop_Terminate()
    On Error Resume Next
    
    i_End = True
End Sub
Public Sub Main()
    On Error Resume Next
    
    ' Load the main window (to get properties).
    Call Load(frmMain)
    
    ' Initialize the string.
    Set RopeString = New clsString
    Call RopeString.SetGravity(9.81)
    Call RopeString.Initialize(100, 30)
    Call RopeString.SetSceneSize((frmMain.ScaleWidth - frmMain.fraDemoControls.Width) - 1, frmMain.ScaleHeight - 1)
    Call RopeString.SetPointPosition(0, ((frmMain.ScaleWidth - frmMain.fraDemoControls.Width) * 0.5), (frmMain.ScaleHeight * 0.75))
    Call RopeString.SetPointIsLocked(0, True)
    
    ' Show the main window.
    Call frmMain.Show(vbModeless)
    
    ' Update the main window's properties.
    'frmMain.lblMaxLength(1).Caption = RopeString.GetMaximumLength() & "px"
    'frmMain.lblPoints(1).Caption = RopeString.GetPointsCount()
    Call frmMain.Controls_SelectPoint(1)
    frmMain.lblPoint(0).ToolTipText = "Current Points Count: " & RopeString.GetPointsCount()
    
    ' Start the loop.
    Call Loop_Start
End Sub
