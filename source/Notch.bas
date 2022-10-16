Attribute VB_Name = "Notch"
'===============================================================================
'   Макрос          : Notch
'   Версия          : 2022.10.16
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk/Notch
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

'===============================================================================
' # константы пользователя

'цвет основных линий
Private Const PrimaryOutlinesColor As String = "RGB255,USER,0,0,255"

'цвет примыкающих линий
Private Const AjacentOutlinesColor As String = "RGB255,USER,255,0,0"

Private Const NotchLength As Double = 3
Private Const NotchWidth As Double = 1
Private Const AdjacencyTolerance As Double = 0.5
Private Const TextLayerName As String = "PartID"
Private Const TextColor As String = "RGB255,USER,255,127,0"
Private Const TextSize As Double = 14

'===============================================================================

Private Type typeNotchParams
    Length As Double
    Width As Double
    Tolerance As Double
End Type

'===============================================================================

Sub Notches()
    MainActivePage False, False
End Sub

Sub NotchesAndText()
    MainActivePage True, False
End Sub

Sub AllDocsNotchesAndExport()
    MainAllDocs False, True
End Sub

Sub AllDocsNotchesExportAndText()
    MainAllDocs True, True
End Sub

'===============================================================================

Private Sub MainActivePage( _
                ByVal AddText As Boolean, _
                ByVal Export As Boolean _
            )
    If RELEASE Then On Error GoTo Catch
    
    If InputData.GetDocumentOrPage.IsError Then Exit Sub
    ActiveDocument.Unit = cdrMillimeter
    
    BoostStart "Notch", RELEASE
    
    MakeNotchesOnActivePage GetNotchParams
    If AddText Then MakeTextOnActivePage
    If Export Then ExportActivePage
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Private Sub MainAllDocs( _
                ByVal AddText As Boolean, _
                ByVal Export As Boolean _
            )
    If RELEASE Then On Error GoTo Catch
    
    If InputData.GetDocumentOrPage.IsError Then Exit Sub
    
    Dim PBar As IProgressBar
    Set PBar = ProgressBar.CreateNumeric(Application.Documents.Count)
    PBar.Caption = "Экспорт"
    PBar.Cancelable = True
    
    Dim Doc As Document
    For Each Doc In Application.Documents
        Doc.Activate
        ActiveDocument.Unit = cdrMillimeter
        BoostStart "Notch", RELEASE
        MakeNotchesOnActivePage GetNotchParams
        If AddText Then MakeTextOnActivePage
        If Export Then ExportActivePage
        BoostFinish
        PBar.Update
        If PBar.Canceled Then Exit For
    Next Doc
    
Finally:
    BoostFinish False
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Private Function GetNotchParams() As typeNotchParams
    GetNotchParams.Length = NotchLength
    GetNotchParams.Width = NotchWidth
    GetNotchParams.Tolerance = AdjacencyTolerance
End Function

Private Function ExportActivePage()
    Dim ExportFilter As ExportFilter
    Set ExportFilter = _
        ActiveDocument.ExportEx( _
            ActiveDocument.FilePath _
          & GetFileNameNoExt(ActiveDocument.FileName) _
          & ".dxf", _
            cdrDXF, cdrCurrentPage _
        )
    With ExportFilter
        .BitmapType = 3 ' FilterDXFLib.dxfBitmapBMP
        .TextAsCurves = False
        .Version = 1 ' FilterDXFLib.dxfVersion2000
        .Units = 3 ' FilterDXFLib.dxfMillimeters
        .FillUnmapped = True
        .FillColor = 0
        .Finish
    End With
End Function

Private Function MakeTextOnActivePage()
    Dim Shapes As ShapeRange
    Set Shapes = ActivePage.Shapes.All
    Dim Text As Shape
    Set Text = CreateOrFindLayer(ActivePage, TextLayerName) _
            .CreateArtisticText(0, 0, GetFileNameNoExt(ActiveDocument.FileName))
    Text.Fill.ApplyUniformFill CreateColor(TextColor)
    Text.Text.Story.Size = TextSize
    Text.CenterX = Shapes.CenterX
    Text.CenterY = Shapes.CenterY
End Function

Private Sub MakeNotchesOnActivePage(ByRef Params As typeNotchParams)
    Dim SourceShapes As ShapeRange
    Set SourceShapes = ActivePage.Shapes.All
    
    Dim PrimaryOutlineColor As Color
    Set PrimaryOutlineColor = CreateColor(PrimaryOutlinesColor)
    Dim AjacentOutlineColor As Color
    Set AjacentOutlineColor = CreateColor(AjacentOutlinesColor)
    
    Dim PrimaryShapes As ShapeRange
    Set PrimaryShapes = _
            FindByCurveAndOutlineColor(SourceShapes, PrimaryOutlineColor)
    Dim AjacentShapes As ShapeRange
    Set AjacentShapes = _
            FindByCurveAndOutlineColor(SourceShapes, AjacentOutlineColor)
    
    Dim PrimaryShape As Shape
    Set PrimaryShape = PrimaryShapes.Combine
    Dim LineWidth As Double
    LineWidth = PrimaryShape.Outline.Width
    
    Dim AjacentShape As Shape
    Dim NotchesAndCuts As New Collection
    For Each AjacentShape In AjacentShapes
        AppendCollection _
            NotchesAndCuts, _
            MakeNotchesOnShape( _
                PrimaryShape, _
                AjacentShape, _
                Params _
            )
    Next AjacentShape
        
    Dim NotchAndCuts As structNotchAndCut
    For Each NotchAndCuts In NotchesAndCuts
        With NotchAndCuts
            If .Success Then
                .Notch.Outline.Color.CopyAssign PrimaryOutlineColor
                .Notch.Outline.Width = LineWidth
                Trim .PrimaryCut, PrimaryShape
                .PrimaryCut.Delete
                .AjacentCut.Delete
            End If
        End With
    Next NotchAndCuts
    
    PrimaryShape.BreakApart
End Sub

Private Function FindByCurveAndOutlineColor( _
                     ByVal Source As ShapeRange, _
                     ByVal OutlineColor As Color _
                 ) As ShapeRange
    Set FindByCurveAndOutlineColor = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Source
        If ShapeHasCurve(Shape) _
       And Shape.Outline.Color.IsSame(OutlineColor) Then
            FindByCurveAndOutlineColor.Add Shape
        End If
    Next Shape
End Function

Private Function MakeNotchesOnShape( _
                     ByVal PrimaryShape As Shape, _
                     ByVal AjacentShape As Shape, _
                     ByRef Params As typeNotchParams _
                 ) As Collection
    Set MakeNotchesOnShape = New Collection
    Dim Node As Node
    For Each Node In AjacentShape.Curve.Nodes
        MakeNotchesOnShape.Add MakeNotch(PrimaryShape, Node, Params)
    Next Node
End Function

Private Function MakeNotch( _
                     ByVal PrimaryShape As Shape, _
                     ByVal AjacentNode As Node, _
                     ByRef Params As typeNotchParams _
                 ) As structNotchAndCut
    Set MakeNotch = New structNotchAndCut
    Dim AjacentSegment As Segment
    Dim Point1 As IPoint
    Dim Point2 As IPoint
    Dim ApexPoint As IPoint
    
    Set AjacentSegment = FindEdgeNodeSegment(AjacentNode)
    
    Dim TempCrossPoints As Collection
    Set MakeNotch.PrimaryCut = _
            PrimaryShape.Layer.CreateEllipse2( _
                AjacentNode.PositionX, _
                AjacentNode.PositionY, _
                Params.Width / 2 _
            )
    Set TempCrossPoints = _
            FindCrossPoints( _
                MakeNotch.PrimaryCut.DisplayCurve, _
                PrimaryShape.Curve _
            )
    TryDeduplicateCrossPoints TempCrossPoints, Params
    If TempCrossPoints.Count < 2 Then
        MakeNotch.PrimaryCut.Delete
        Exit Function
    End If
    Set Point1 = TempCrossPoints(1)
    Set Point2 = TempCrossPoints(2)
    
    Set MakeNotch.AjacentSubPath = AjacentSegment.SubPath
    Set MakeNotch.AjacentCut = _
            PrimaryShape.Layer.CreateEllipse2( _
                AjacentNode.PositionX, _
                AjacentNode.PositionY, _
                Params.Length _
            )
    Set TempCrossPoints = _
            FindCrossPoints( _
                MakeNotch.AjacentCut.DisplayCurve, _
                MakeNotch.AjacentSubPath _
            )
    If TempCrossPoints.Count < 1 Then
        MakeNotch.PrimaryCut.Delete
        MakeNotch.AjacentCut.Delete
        Exit Function
    End If
    Set ApexPoint = TempCrossPoints(1)

    Set MakeNotch.Notch = DrawNotch(Point1, Point2, ApexPoint)
    
    MakeNotch.Success = True
End Function

Private Sub TryDeduplicateCrossPoints( _
                ByRef ioCrossPoints As Collection, _
                ByRef Params As typeNotchParams _
            )
    If ioCrossPoints.Count < 3 Then Exit Sub
    Dim NewPoints As New Collection
    Do While ioCrossPoints.Count > 0
        If Not PointDuplicates(ioCrossPoints(1), ioCrossPoints, Params) Then
            NewPoints.Add ioCrossPoints(1)
        End If
        ioCrossPoints.Remove 1
    Loop
    Set ioCrossPoints = NewPoints
End Sub

Private Function PointDuplicates( _
                ByVal CrossPoint As IPoint, _
                ByVal CrossPoints As Collection, _
                ByRef Params As typeNotchParams _
            ) As Boolean
    If CrossPoints.Count < 2 Then Exit Function
    Dim Point As IPoint
    For Each Point In CrossPoints
        If Not Point Is CrossPoint Then
            If Point.GetDistanceFrom(CrossPoint) < Params.Tolerance Then
                PointDuplicates = True
                Exit Function
            End If
        End If
    Next Point
End Function

Private Function DrawNotch( _
                     ByVal Point1 As IPoint, _
                     ByVal Point2 As IPoint, _
                     ByVal ApexPoint As IPoint _
                 ) As Shape
    Dim SubPath As SubPath
    Set SubPath = _
            ActiveDocument.CreateCurve.CreateSubPath( _
                Point1.x, Point1.y _
            )
    SubPath.AppendCurveSegment ApexPoint.x, ApexPoint.y
    SubPath.AppendCurveSegment Point2.x, Point2.y
    Set DrawNotch = ActiveLayer.CreateCurve(SubPath.Parent)
End Function

Private Function MakeNodeAtEdgeNode( _
                     ByVal Segment As Segment, _
                     ByVal Offset As Double _
                 ) As Node
    If Segment.EndNode.IsEnding Then Offset = Segment.Length - Offset
    Set MakeNodeAtEdgeNode = _
            Segment.AddNodeAt(Offset, cdrAbsoluteSegmentOffset)
End Function

Private Function FindNodesOnShape( _
                     ByVal Nodes As Nodes, _
                     ByVal Shape As Shape, _
                     Optional ByVal Tolerance As Double = 0.001 _
                 ) As Collection
    Set FindNodesOnShape = New Collection
    Dim Offset As Double
    Dim Seg As Segment
    Dim Node As Node
    For Each Node In Nodes
        Set Seg = _
                Shape.Curve.FindSegmentAtPoint( _
                    Node.PositionX, _
                    Node.PositionY, _
                    Offset, _
                    Tolerance _
                )
        If Not Seg Is Nothing Then
            FindNodesOnShape.Add Node
        End If
    Next Node
End Function

Private Function FindCrossPoints( _
                     ByVal CurveOrSubPath1 As Variant, _
                     ByVal CurveOrSubPath2 As Variant _
                 ) As Collection
    Set FindCrossPoints = New Collection
    Dim Seg1 As Segment
    Dim Seg2 As Segment
    Dim Points As CrossPoints
    For Each Seg1 In CurveOrSubPath1.Segments
        For Each Seg2 In CurveOrSubPath2.Segments
            Set Points = Seg1.GetIntersections(Seg2)
            If Points.Count > 0 Then _
                AddCrossPointsToCollection FindCrossPoints, Points
        Next Seg2
    Next Seg1
End Function

Private Sub AddCrossPointsToCollection( _
                ByVal ioCollection As Collection, _
                ByVal CrossPoints As CrossPoints _
            )
    Dim Point As CrossPoint
    For Each Point In CrossPoints
        ioCollection.Add FreePoint.Create(Point.PositionX, Point.PositionY)
    Next Point
End Sub

Private Function FindEdgeNodeSegment(ByVal Node As Node) As Segment
    If Not Node.Segment Is Nothing Then
        Set FindEdgeNodeSegment = Node.Segment
        Exit Function
    End If
    If Not Node.NextSegment Is Nothing Then
        Set FindEdgeNodeSegment = Node.NextSegment
        Exit Function
    End If
    If Not Node.PrevSegment Is Nothing Then
        Set FindEdgeNodeSegment = Node.PrevSegment
        Exit Function
    End If
End Function
