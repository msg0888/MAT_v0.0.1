Attribute VB_Name = "Globals"
Public options As New ModelLoadOptions
 
Public LoadCatMap As New Scripting.Dictionary
Public DirMap As New Scripting.Dictionary
Public MaterialTypeMap As New Scripting.Dictionary


Public Function cNode() As Node
  Set cNode = New Node
End Function

Public Function cShape() As Shape
  Set cShape = New Shape
End Function

Public Function cPlate() As Plate
  Set cPlate = New Plate
End Function

Public Function cBasicLoadCase() As BasicLoadCase
  Set cBasicLoadCase = New BasicLoadCase
End Function

Public Function cPointLoad() As PointLoad
  Set cPointLoad = New PointLoad
End Function

Public Function cDistributedLoad() As DistributedLoad
  Set cDistributedLoad = New DistributedLoad
End Function

Public Function cAreaLoad() As AreaLoad
  Set cAreaLoad = New AreaLoad
End Function

Public Function cSurfaceLoad() As SurfaceLoad
  Set cSurfaceLoad = New SurfaceLoad
End Function

Public Function cLoadCombination() As LoadCombination
  Set cLoadCombination = New LoadCombination
End Function
