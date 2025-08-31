Attribute VB_Name = "mdMain"
' mdMain - Run the simulation demo
' Purpose: Starts the simulation


Option Explicit

Public Sub RunSimulation()
    Dim config As New clsSimulationConfig
    
    ' (optional) tweak defaults:
    'cfg.Width = 80: cfg.Height = 50
    'cfg.AliveProb = 0.25
    'cfg.PaintEvery = 1
    'cfg.WrapEdges = True

    Dim simulation As New clsSimulation
  
    simulation.Init config, New clsGameOfLife, New clsGridRenderer
    simulation.Run topLeft:=ActiveSheet.Cells(1, 1)
    
    Set simulation = Nothing
    Set config = Nothing

End Sub

