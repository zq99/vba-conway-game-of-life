VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSimulationConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsSimulationConfig.cls — Run settings
' Purpose: Validated, (effectively) immutable settings for a simulation run:
'          board size, initial live-cell probability, wrap behavior, max iters,
'          and render cadence. Provides Validate/Freeze to guard correctness.

Option Explicit

Private mlngWidth As Long, mlngHeight As Long
Private mdblAliveProb As Double
Private mlngMaxIters As Long
Private mblnWrapEdges As Boolean
Private mlngPaintEvery As Long
Private mblnFrozen As Boolean

Private Sub Class_Initialize()
    mlngWidth = 60: mlngHeight = 40
    mdblAliveProb = 0.22
    mlngMaxIters = 10000
    mblnWrapEdges = False
    mlngPaintEvery = 1
End Sub

Public Property Get width() As Long: width = mlngWidth: End Property
Public Property Let width(ByVal v As Long): GuardNotFrozen: mlngWidth = v: End Property

Public Property Get height() As Long: height = mlngHeight: End Property
Public Property Let height(ByVal v As Long): GuardNotFrozen: mlngHeight = v: End Property

Public Property Get aliveProb() As Double: aliveProb = mdblAliveProb: End Property
Public Property Let aliveProb(ByVal v As Double): GuardNotFrozen: mdblAliveProb = v: End Property

Public Property Get MaxIters() As Long: MaxIters = mlngMaxIters: End Property
Public Property Let MaxIters(ByVal v As Long): GuardNotFrozen: mlngMaxIters = v: End Property

Public Property Get wrapEdges() As Boolean: wrapEdges = mblnWrapEdges: End Property
Public Property Let wrapEdges(ByVal v As Boolean): GuardNotFrozen: mblnWrapEdges = v: End Property

Public Property Get PaintEvery() As Long: PaintEvery = mlngPaintEvery: End Property
Public Property Let PaintEvery(ByVal v As Long): GuardNotFrozen: mlngPaintEvery = v: End Property

Public Property Get IsFrozen() As Boolean: IsFrozen = mblnFrozen: End Property

Public Sub Validate()
    If mlngWidth <= 0 Or mlngHeight <= 0 Then Err.Raise 5, , "Width/Height must be > 0."
    If mdblAliveProb < 0# Or mdblAliveProb > 1# Then Err.Raise 5, , "AliveProb must be 0..1."
    If mlngMaxIters < 1 Then Err.Raise 5, , "MaxIters must be >= 1."
    If mlngPaintEvery < 1 Then mlngPaintEvery = 1
End Sub

Public Sub Freeze(): mblnFrozen = True: End Sub

Private Sub GuardNotFrozen()
    If mblnFrozen Then Err.Raise 5, , "Config is frozen for this run."
End Sub

