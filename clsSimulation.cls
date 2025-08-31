VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsSimulation.cls — Orchestrator
' Purpose: Wires together config, model, and renderer; manages Excel host flags
'          (ScreenUpdating, Calculation, Events), runs the simulation loop, and
'          handles cancellation and clean teardown.


Option Explicit

Private mConfig As clsSimulationConfig
Private mAutomaton As ICellularAutomaton
Private mRenderer As IRenderer
Private mEnmOldCancel As XlEnableCancelKey
Private mEnmPrevCalc As XlCalculation
Private mblnWasUpdating As Boolean
Private mblnPrevEvents As Boolean
Private mblnPrevPageBreaks As Boolean
Private mblnCancelled As Boolean
Private Const CINT_ERROR_USER_INTERRUPT As Integer = 18

Public Sub Init(Optional ByVal cfg As clsSimulationConfig, _
                Optional ByVal automaton As ICellularAutomaton, _
                Optional ByVal renderer As IRenderer)
                
    If cfg Is Nothing Then Set mConfig = New clsSimulationConfig Else Set mConfig = cfg
    If automaton Is Nothing Then Err.Raise 5, , "Model dependency not provided."
    If renderer Is Nothing Then Err.Raise 5, , "Renderer dependency not provided."
    Set mAutomaton = automaton: Set mRenderer = renderer
    mConfig.Validate: mConfig.Freeze

End Sub

Public Sub Run(ByVal topLeft As Range)
    On Error GoTo ErrHandler
    
    
    Dim ws As Worksheet
    
    Set ws = topLeft.Parent

    SetupHost ws
    ws.Cells.Clear

    mAutomaton.Init mConfig.width, mConfig.height, mConfig.wrapEdges
    mAutomaton.SeedRandom mConfig.aliveProb
    mRenderer.Init ws, topLeft, mConfig.height, mConfig.width
    mRenderer.Clear

    Dim it As Long, snap As Variant
    For it = 1 To mConfig.MaxIters
        DoEvents: If mblnCancelled Then Exit For

        If (it Mod mConfig.PaintEvery) = 0 Then
            snap = mAutomaton.Snapshot

            ' paint with SU off inside renderer, then pulse once to show the frame
            mRenderer.Render snap
            Application.ScreenUpdating = True  ' flush once per frame
            DoEvents
            Application.ScreenUpdating = False
        End If

        mAutomaton.StepOne
    Next it

CleanExit:
    Set ws = Nothing
    TeardownHost
    Exit Sub

ErrHandler:
    Dim n As Long, d As String, s As String
    n = Err.Number: d = Err.Description: s = Err.Source
    TeardownHost
    If n = CINT_ERROR_USER_INTERRUPT Then Exit Sub
    Err.Raise n, s, d
End Sub

Public Sub Cancel(): mblnCancelled = True: End Sub

Private Sub SetupHost(ByVal ws As Worksheet)
    mblnWasUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    mEnmOldCancel = Application.EnableCancelKey
    Application.EnableCancelKey = xlErrorHandler

    mEnmPrevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    mblnPrevEvents = Application.EnableEvents
    Application.EnableEvents = False

    mblnPrevPageBreaks = ws.DisplayPageBreaks
    ws.DisplayPageBreaks = False
End Sub

Private Sub TeardownHost()
    Application.ScreenUpdating = mblnWasUpdating
    Application.EnableCancelKey = mEnmOldCancel
    Application.Calculation = mEnmPrevCalc
    Application.EnableEvents = mblnPrevEvents
    On Error Resume Next
    ActiveSheet.DisplayPageBreaks = mblnPrevPageBreaks
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Set mConfig = Nothing
    Set mAutomaton = Nothing
    Set mRenderer = Nothing
End Sub
