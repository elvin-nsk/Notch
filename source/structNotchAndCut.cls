VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "structNotchAndCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Struct
Option Explicit

'===============================================================================

Public Success As Boolean
Public Notch As Shape
Public PrimaryCut As Shape
Public AjacentCut As Shape
Public AjacentSubPath As SubPath
