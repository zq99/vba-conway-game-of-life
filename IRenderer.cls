VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub Init(ByVal ws As Worksheet, ByVal topLeft As Range, _
                ByVal height As Long, ByVal width As Long): End Sub
Public Sub Clear(): End Sub
Public Sub Render(ByVal grid As Variant): End Sub


