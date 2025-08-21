VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AreaDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------
' Class Module: AreaDef
' Description: Represents a movable area to be placed in a zone.
'--------------------------------------------
Option Explicit

Public Row As Long
Public Name As String
Public workload As Double
Public width As Double
Public height As Double
Public centerX As Double
Public centerY As Double
