Attribute VB_Name = "AutoAnimate"
'===========================================================================
' AutoAnimate.bas
'   Automatically add entrance animations to the shapes on a PowerPoint slide.
'
'   Logic mirrors the Python tool (auto_animate.py):
'     - Regular shapes     -> Fade entrance
'     - Text boxes         -> Wipe from top (reveal top-to-bottom)
'     - Arrows / connectors-> Wipe in the direction the arrow points
'     - Flow ordering      : if arrow goes A -> B, reveal A, arrow+B on clicks
'     - Orphan text/shapes : above the flow fire first, below fire last
'     - Slides already animated are skipped
'     - Slides without "[auto-animate]" in notes are skipped
'
'   HOW TO INSTALL
'     1.  Open your .pptx in PowerPoint.
'     2.  Press Alt+F11 to open the VBA editor.
'     3.  File -> Import File... -> select AutoAnimate.bas.
'     4.  Close the editor.  Enable macros for this session when prompted.
'
'   HOW TO RUN
'     -  Put "[auto-animate]" into the speaker-notes pane of every slide you
'        want animated, then press Alt+F8, pick "AutoAnimate", and Run.
'     -  Or run "AutoAnimateAll" to animate every slide regardless of marker.
'     -  "AutoAnimateCurrent" just animates the currently selected slide.
'
'===========================================================================

Option Explicit

Private Const MARKER As String = "[auto-animate]"

' -------- edge (arrow-direction) constants ---------------------------------
Private Const DIR_RIGHT As Long = 1
Private Const DIR_LEFT  As Long = 2
Private Const DIR_DOWN  As Long = 3
Private Const DIR_UP    As Long = 4

' -------- shape kinds ------------------------------------------------------
Private Const KIND_TEXT       As String = "text"
Private Const KIND_SHAPE      As String = "shape"
Private Const KIND_CONNECTOR  As String = "connector"
Private Const KIND_BLOCK_ARR  As String = "block_arrow"

' -------- small record of per-shape info -----------------------------------
Private Type ShapeRec
    idx      As Long      ' Slide.Shapes index
    kind     As String
    leftX    As Single
    topY     As Single
    width_   As Single
    height_  As Single
    tailX    As Single
    tailY    As Single
    headX    As Single
    headY    As Single
    motion   As Long      ' DIR_*  -- only meaningful for arrows
End Type

'---------------------------------------------------------------------------
'  PUBLIC ENTRY POINTS
'---------------------------------------------------------------------------
Public Sub AutoAnimate()
    Dim sl As slide, n As Long
    For Each sl In ActivePresentation.Slides
        If Not SlideHasMarker(sl) Then
            Debug.Print "slide " & sl.SlideIndex & ": skipped (no marker)"
        ElseIf SlideHasAnimations(sl) Then
            Debug.Print "slide " & sl.SlideIndex & ": skipped (already animated)"
        Else
            Debug.Print "slide " & sl.SlideIndex & ": processing"
            AnimateSlide sl
            n = n + 1
        End If
    Next sl
    MsgBox "AutoAnimate done: " & n & " slide(s) animated.", vbInformation
End Sub

Public Sub AutoAnimateAll()
    Dim sl As slide, n As Long
    For Each sl In ActivePresentation.Slides
        If Not SlideHasAnimations(sl) Then
            AnimateSlide sl
            n = n + 1
        End If
    Next sl
    MsgBox "AutoAnimateAll done: " & n & " slide(s) animated.", vbInformation
End Sub

Public Sub AutoAnimateCurrent()
    Dim sl As slide
    Set sl = Application.ActiveWindow.View.slide
    If SlideHasAnimations(sl) Then
        MsgBox "Slide already has animations.", vbExclamation
        Exit Sub
    End If
    AnimateSlide sl
End Sub

'---------------------------------------------------------------------------
'  MARKER / TIMING HELPERS
'---------------------------------------------------------------------------
Private Function SlideHasMarker(sl As slide) As Boolean
    On Error Resume Next
    Dim t As String
    Dim ph As Shape
    If Not sl.HasNotesPage Then Exit Function
    For Each ph In sl.NotesPage.Shapes
        If ph.HasTextFrame Then
            If ph.TextFrame.HasText Then
                t = t & " " & ph.TextFrame.TextRange.Text
            End If
        End If
    Next ph
    SlideHasMarker = (InStr(1, t, MARKER) > 0)
End Function

Private Function SlideHasAnimations(sl As slide) As Boolean
    SlideHasAnimations = (sl.TimeLine.MainSequence.Count > 0)
End Function

'---------------------------------------------------------------------------
'  CLASSIFY ONE SHAPE
'---------------------------------------------------------------------------
Private Function ClassifyShape(s As Shape, rec As ShapeRec) As Boolean
    rec.leftX = s.Left:   rec.topY = s.Top
    rec.width_ = s.width: rec.height_ = s.height

    ' Connector (line)
    If s.Type = msoLine Or s.connector = msoTrue Then
        rec.kind = KIND_CONNECTOR
        ComputeConnectorEndpoints s, rec
        Exit Function  '
    End If

    ' Block-arrow auto shape
    If s.Type = msoAutoShape Then
        Dim dir_ As Long
        dir_ = BlockArrowDirection(s.AutoShapeType)
        If dir_ <> 0 Then
            rec.kind = KIND_BLOCK_ARR
            ' Apply horizontal/vertical flip and rotation to the base direction.
            If s.HorizontalFlip = msoTrue Then dir_ = FlipDirH(dir_)
            If s.VerticalFlip = msoTrue Then dir_ = FlipDirV(dir_)
            rec.motion = RotateDir(dir_, s.Rotation)
            ComputeBlockArrowEndpoints s, rec
            ClassifyShape = True
            Exit Function
        End If
    End If

    ' Text box
    If s.Type = msoTextBox Then
        rec.kind = KIND_TEXT
        ClassifyShape = True
        Exit Function
    End If

    ' Default: regular shape
    rec.kind = KIND_SHAPE
    ClassifyShape = True
End Function

Private Function BlockArrowDirection(t As MsoAutoShapeType) As Long
    Select Case t
        Case msoShapeRightArrow, msoShapeStripedRightArrow, _
             msoShapeNotchedRightArrow, msoShapeCurvedRightArrow
            BlockArrowDirection = DIR_RIGHT
        Case msoShapeLeftArrow, msoShapeCurvedLeftArrow
            BlockArrowDirection = DIR_LEFT
        Case msoShapeUpArrow, msoShapeBentUpArrow, msoShapeCurvedUpArrow
            BlockArrowDirection = DIR_UP
        Case msoShapeDownArrow, msoShapeCurvedDownArrow
            BlockArrowDirection = DIR_DOWN
        Case Else
            BlockArrowDirection = 0
    End Select
End Function

Private Function FlipDirH(d As Long) As Long
    Select Case d
        Case DIR_RIGHT: FlipDirH = DIR_LEFT
        Case DIR_LEFT:  FlipDirH = DIR_RIGHT
        Case Else:      FlipDirH = d
    End Select
End Function

Private Function FlipDirV(d As Long) As Long
    Select Case d
        Case DIR_DOWN: FlipDirV = DIR_UP
        Case DIR_UP:   FlipDirV = DIR_DOWN
        Case Else:     FlipDirV = d
    End Select
End Function

Private Function RotateDir(d As Long, degreesCW As Single) As Long
    ' Snap rotation to nearest 90° step and apply that many CW quarter turns.
    Dim q As Long
    q = ((CLng(Int((CDbl(degreesCW) + 45) / 90)) Mod 4) + 4) Mod 4
    Dim i As Long, r As Long: r = d
    For i = 1 To q
        ' rotate one quarter turn CW: RIGHT->DOWN->LEFT->UP->RIGHT
        Select Case r
            Case DIR_RIGHT: r = DIR_DOWN
            Case DIR_DOWN:  r = DIR_LEFT
            Case DIR_LEFT:  r = DIR_UP
            Case DIR_UP:    r = DIR_RIGHT
        End Select
    Next i
    RotateDir = r
End Function

Private Sub ComputeConnectorEndpoints(s As Shape, rec As ShapeRec)
    ' Use bounding box corners modified by HorizontalFlip/VerticalFlip.
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single
    x1 = s.Left: y1 = s.Top
    x2 = s.Left + s.width: y2 = s.Top + s.height
    If s.HorizontalFlip = msoTrue Then
        Dim tmpX As Single: tmpX = x1: x1 = x2: x2 = tmpX
    End If
    If s.VerticalFlip = msoTrue Then
        Dim tmpY As Single: tmpY = y1: y1 = y2: y2 = tmpY
    End If
    ' If the arrowhead is on the BEGIN end (x1,y1) and not on the END end
    ' (x2,y2), the line is visually reversed — swap so (x2,y2) is the pointy end.
    Dim beginArr As Boolean, endArr As Boolean
    On Error Resume Next
    beginArr = (s.Line.BeginArrowheadStyle <> msoArrowheadNone)
    endArr   = (s.Line.EndArrowheadStyle   <> msoArrowheadNone)
    On Error GoTo 0
    If beginArr And Not endArr Then
        Dim tX As Single, tY As Single
        tX = x1: x1 = x2: x2 = tX
        tY = y1: y1 = y2: y2 = tY
    End If
    rec.tailX = x1: rec.tailY = y1
    rec.headX = x2: rec.headY = y2
    rec.motion = DirectionOf(x2 - x1, y2 - y1)
End Sub

Private Sub ComputeBlockArrowEndpoints(s As Shape, rec As ShapeRec)
    ' Uses the final motion direction (post-flip, post-rotate) to place tail/head
    ' on the bbox edges.  The bbox we have is the un-rotated one, but its centre
    ' doesn't move with rotation, so this is a reasonable approximation for
    ' flow-endpoint matching.
    Dim cx As Single, cy As Single
    cx = s.Left + s.width / 2
    cy = s.Top + s.height / 2
    Select Case rec.motion
        Case DIR_RIGHT
            rec.tailX = s.Left: rec.tailY = cy
            rec.headX = s.Left + s.width: rec.headY = cy
        Case DIR_LEFT
            rec.tailX = s.Left + s.width: rec.tailY = cy
            rec.headX = s.Left: rec.headY = cy
        Case DIR_DOWN
            rec.tailX = cx: rec.tailY = s.Top
            rec.headX = cx: rec.headY = s.Top + s.height
        Case DIR_UP
            rec.tailX = cx: rec.tailY = s.Top + s.height
            rec.headX = cx: rec.headY = s.Top
    End Select
End Sub

Private Function DirectionOf(dx As Single, dy As Single) As Long
    If Abs(dx) >= Abs(dy) Then
        If dx >= 0 Then DirectionOf = DIR_RIGHT Else DirectionOf = DIR_LEFT
    Else
        If dy >= 0 Then DirectionOf = DIR_DOWN Else DirectionOf = DIR_UP
    End If
End Function

'---------------------------------------------------------------------------
'  SPATIAL-FLOW INFERENCE
'---------------------------------------------------------------------------
Private Function NearestShape(pt_x As Single, pt_y As Single, _
                              recs() As ShapeRec, nRecs As Long, _
                              excludeKind1 As String, excludeKind2 As String) As Long
    ' 1. any shape whose bbox contains the point
    Dim i As Long, best As Long, bestArea As Single, bestDist As Single
    best = -1
    For i = 0 To nRecs - 1
        If recs(i).kind <> excludeKind1 And recs(i).kind <> excludeKind2 Then
            If PointInRect(pt_x, pt_y, recs(i)) Then
                Dim area As Single
                area = recs(i).width_ * recs(i).height_
                If best = -1 Or area < bestArea Then
                    best = i
                    bestArea = area
                End If
            End If
        End If
    Next i
    If best >= 0 Then
        NearestShape = best
        Exit Function
    End If

    ' 2. nearest centre, with a max-distance cutoff (~2.3 inches = 165 points)
    Const MAX_D As Single = 165
    best = -1
    For i = 0 To nRecs - 1
        If recs(i).kind <> excludeKind1 And recs(i).kind <> excludeKind2 Then
            Dim d As Single
            d = DistToCenter(pt_x, pt_y, recs(i))
            If d < MAX_D Then
                If best = -1 Or d < bestDist Then
                    best = i
                    bestDist = d
                End If
            End If
        End If
    Next i
    NearestShape = best
End Function

Private Function PointInRect(px As Single, py As Single, r As ShapeRec) As Boolean
    Const PAD As Single = 4 ' points
    PointInRect = (px >= r.leftX - PAD And px <= r.leftX + r.width_ + PAD And _
                   py >= r.topY - PAD And py <= r.topY + r.height_ + PAD)
End Function

Private Function DistToCenter(px As Single, py As Single, r As ShapeRec) As Single
    Dim cx As Single, cy As Single
    cx = r.leftX + r.width_ / 2
    cy = r.topY + r.height_ / 2
    DistToCenter = Sqr((px - cx) ^ 2 + (py - cy) ^ 2)
End Function

'---------------------------------------------------------------------------
'  MAIN SLIDE-LEVEL ROUTINE
'---------------------------------------------------------------------------
Private Sub AnimateSlide(sl As slide)
    Dim recs() As ShapeRec, nRecs As Long
    ReDim recs(0 To sl.Shapes.Count - 1)
    Dim i As Long
    For i = 1 To sl.Shapes.Count
        Dim r As ShapeRec
        If ClassifyShape(sl.Shapes(i), r) Or True Then
            r.idx = i
            recs(nRecs) = r
            nRecs = nRecs + 1
        End If
    Next i
    If nRecs = 0 Then Exit Sub
    ReDim Preserve recs(0 To nRecs - 1)

    ' Build arrows -> (src, dst) using NearestShape
    Dim srcIdx() As Long, dstIdx() As Long
    ReDim srcIdx(0 To nRecs - 1): ReDim dstIdx(0 To nRecs - 1)
    For i = 0 To nRecs - 1
        srcIdx(i) = -1: dstIdx(i) = -1
        If recs(i).kind = KIND_CONNECTOR Or recs(i).kind = KIND_BLOCK_ARR Then
            srcIdx(i) = NearestShape(recs(i).tailX, recs(i).tailY, recs, nRecs, _
                                     KIND_CONNECTOR, KIND_BLOCK_ARR)
            dstIdx(i) = NearestShape(recs(i).headX, recs(i).headY, recs, nRecs, _
                                     KIND_CONNECTOR, KIND_BLOCK_ARR)
        End If
    Next i

    ' Build adjacency counts
    Dim incoming() As Long: ReDim incoming(0 To nRecs - 1)
    For i = 0 To nRecs - 1
        If srcIdx(i) >= 0 And dstIdx(i) >= 0 Then incoming(dstIdx(i)) = incoming(dstIdx(i)) + 1
    Next i

    ' Determine reading order (top, then left) over non-arrows
    Dim readingOrder() As Long, nReading As Long
    ReDim readingOrder(0 To nRecs - 1)
    For i = 0 To nRecs - 1
        If recs(i).kind <> KIND_CONNECTOR And recs(i).kind <> KIND_BLOCK_ARR Then
            readingOrder(nReading) = i
            nReading = nReading + 1
        End If
    Next i
    If nReading = 0 Then Exit Sub
    ReDim Preserve readingOrder(0 To nReading - 1)
    SortByTopLeft recs, readingOrder, nReading

    ' BFS — flow order into click groups
    Dim visitedShape() As Boolean, visitedArrow() As Boolean, processedOut() As Boolean
    ReDim visitedShape(0 To nRecs - 1)
    ReDim visitedArrow(0 To nRecs - 1)
    ReDim processedOut(0 To nRecs - 1)

    ' groups(clickNum, memberSlot) = recs-index; members() holds count
    Dim groups() As Long, groupSizes() As Long, nGroups As Long
    ReDim groups(0 To 255, 0 To 15): ReDim groupSizes(0 To 255)
    nGroups = 0

    ' Phase 1: BFS from sources (incoming=0, has outgoing)
    Dim ro As Long
    For ro = 0 To nReading - 1
        Dim startI As Long: startI = readingOrder(ro)
        If Not visitedShape(startI) And incoming(startI) = 0 And HasOutgoing(startI, srcIdx, nRecs) Then
            BFSFromNode startI, recs, srcIdx, dstIdx, nRecs, _
                        visitedShape, visitedArrow, processedOut, _
                        groups, groupSizes, nGroups
        End If
    Next ro

    ' Phase 2: any still-unvisited node that has any in/out flow edges
    For ro = 0 To nReading - 1
        Dim idx2 As Long: idx2 = readingOrder(ro)
        If Not visitedShape(idx2) Then
            If incoming(idx2) > 0 Or HasOutgoing(idx2, srcIdx, nRecs) Then
                BFSFromNode idx2, recs, srcIdx, dstIdx, nRecs, _
                            visitedShape, visitedArrow, processedOut, _
                            groups, groupSizes, nGroups
            End If
        End If
    Next ro

    ' Phase 3: leftover arrows
    For i = 0 To nRecs - 1
        If (recs(i).kind = KIND_CONNECTOR Or recs(i).kind = KIND_BLOCK_ARR) _
           And Not visitedArrow(i) Then
            groups(nGroups, 0) = i: groupSizes(nGroups) = 1
            nGroups = nGroups + 1
            visitedArrow(i) = True
        End If
    Next i

    ' Phase 4: orphan non-arrows. Split by top above/below flow's topmost non-arrow.
    Dim divider As Single: divider = 1E+20
    Dim g As Long, m As Long
    For g = 0 To nGroups - 1
        For m = 0 To groupSizes(g) - 1
            Dim gi As Long: gi = groups(g, m)
            If recs(gi).kind <> KIND_CONNECTOR And recs(gi).kind <> KIND_BLOCK_ARR Then
                If recs(gi).topY < divider Then divider = recs(gi).topY
            End If
        Next m
    Next g
    If divider = 1E+20 Then divider = 0

    ' Split orphans
    Dim preIdx() As Long, postIdx() As Long, nPre As Long, nPost As Long
    ReDim preIdx(0 To nReading - 1): ReDim postIdx(0 To nReading - 1)
    For ro = 0 To nReading - 1
        Dim oidx As Long: oidx = readingOrder(ro)
        If Not visitedShape(oidx) Then
            If recs(oidx).topY < divider Then
                preIdx(nPre) = oidx: nPre = nPre + 1
            Else
                postIdx(nPost) = oidx: nPost = nPost + 1
            End If
        End If
    Next ro

    ' Materialize final click-group order into the timeline
    Dim click As Long
    ' pre-orphans first
    For click = 0 To nPre - 1
        ApplyClickGroup sl, recs, Array(preIdx(click)), 1
    Next click
    ' flow groups
    For click = 0 To nGroups - 1
        Dim members() As Long, sz As Long
        sz = groupSizes(click)
        ReDim members(0 To sz - 1)
        For m = 0 To sz - 1
            members(m) = groups(click, m)
        Next m
        ApplyClickGroup sl, recs, members, sz
    Next click
    ' post-orphans last
    For click = 0 To nPost - 1
        ApplyClickGroup sl, recs, Array(postIdx(click)), 1
    Next click
End Sub

Private Function HasOutgoing(nodeIdx As Long, srcIdx() As Long, nRecs As Long) As Boolean
    Dim i As Long
    For i = 0 To nRecs - 1
        If srcIdx(i) = nodeIdx Then
            HasOutgoing = True
            Exit Function
        End If
    Next i
End Function

Private Sub BFSFromNode(startIdx As Long, _
                        recs() As ShapeRec, srcIdx() As Long, dstIdx() As Long, nRecs As Long, _
                        ByRef visitedShape() As Boolean, _
                        ByRef visitedArrow() As Boolean, _
                        ByRef processedOut() As Boolean, _
                        ByRef groups() As Long, _
                        ByRef groupSizes() As Long, _
                        ByRef nGroups As Long)
    Dim queue() As Long, head As Long, tail As Long
    ReDim queue(0 To nRecs - 1)
    queue(0) = startIdx: head = 0: tail = 1

    Do While head < tail
        Dim node As Long: node = queue(head): head = head + 1
        If Not visitedShape(node) Then
            visitedShape(node) = True
            groups(nGroups, 0) = node: groupSizes(nGroups) = 1
            nGroups = nGroups + 1
        End If
        If processedOut(node) Then GoTo nextq
        processedOut(node) = True

        ' Collect this node's outgoing arrows, sorted by destination reading
        ' order (top-to-bottom, then left-to-right).
        Dim outArrows() As Long, nOut As Long
        ReDim outArrows(0 To nRecs - 1)
        nOut = 0
        Dim a As Long
        For a = 0 To nRecs - 1
            If srcIdx(a) = node And Not visitedArrow(a) Then
                outArrows(nOut) = a
                nOut = nOut + 1
            End If
        Next a
        ' Selection sort on outArrows by destination's (top, left)
        Dim i2 As Long, j2 As Long, bestI As Long
        For i2 = 0 To nOut - 2
            bestI = i2
            For j2 = i2 + 1 To nOut - 1
                If DstPrecedes(outArrows(j2), outArrows(bestI), dstIdx, recs) Then
                    bestI = j2
                End If
            Next j2
            If bestI <> i2 Then
                Dim tmp As Long: tmp = outArrows(i2)
                outArrows(i2) = outArrows(bestI): outArrows(bestI) = tmp
            End If
        Next i2

        Dim k As Long
        For k = 0 To nOut - 1
            a = outArrows(k)
            visitedArrow(a) = True
            Dim dst As Long: dst = dstIdx(a)
            Dim size As Long: size = 1
            groups(nGroups, 0) = a
            If dst >= 0 And Not visitedShape(dst) Then
                groups(nGroups, 1) = dst
                size = 2
                visitedShape(dst) = True
            End If
            groupSizes(nGroups) = size
            nGroups = nGroups + 1
            If dst >= 0 Then
                queue(tail) = dst: tail = tail + 1
            End If
        Next k
nextq:
    Loop
End Sub

Private Function DstPrecedes(aIdx As Long, bIdx As Long, _
                             dstIdx() As Long, recs() As ShapeRec) As Boolean
    ' True iff arrow aIdx's destination should animate before bIdx's destination.
    Dim da As Long: da = dstIdx(aIdx)
    Dim db As Long: db = dstIdx(bIdx)
    If da < 0 And db < 0 Then DstPrecedes = False: Exit Function
    If da < 0 Then DstPrecedes = False: Exit Function
    If db < 0 Then DstPrecedes = True:  Exit Function
    If recs(da).topY < recs(db).topY Then DstPrecedes = True: Exit Function
    If recs(da).topY > recs(db).topY Then DstPrecedes = False: Exit Function
    DstPrecedes = (recs(da).leftX < recs(db).leftX)
End Function

Private Sub SortByTopLeft(recs() As ShapeRec, idxs() As Long, n As Long)
    ' Simple insertion sort (n is small).
    Dim i As Long, j As Long, key As Long
    For i = 1 To n - 1
        key = idxs(i): j = i - 1
        Do While j >= 0
            If recs(idxs(j)).topY < recs(key).topY Then Exit Do
            If recs(idxs(j)).topY = recs(key).topY And recs(idxs(j)).leftX <= recs(key).leftX Then Exit Do
            idxs(j + 1) = idxs(j)
            j = j - 1
        Loop
        idxs(j + 1) = key
    Next i
End Sub

'---------------------------------------------------------------------------
'  APPLY A CLICK GROUP  (one click -> one or more effects running together)
'---------------------------------------------------------------------------
Private Sub ApplyClickGroup(sl As slide, recs() As ShapeRec, members As Variant, sz As Long)
    Dim i As Long
    For i = 0 To sz - 1
        Dim ridx As Long
        If IsArray(members) Then ridx = members(i) Else ridx = members
        Dim sp As Shape: Set sp = sl.Shapes(recs(ridx).idx)
        Dim trig As MsoAnimTriggerType
        If i = 0 Then trig = msoAnimTriggerOnPageClick Else trig = msoAnimTriggerWithPrevious
        AddOneEffect sl, sp, recs(ridx), trig
    Next i
End Sub

Private Sub AddOneEffect(sl As slide, sp As Shape, rec As ShapeRec, trig As MsoAnimTriggerType)
    Dim fx As Effect
    Select Case rec.kind
        Case KIND_CONNECTOR, KIND_BLOCK_ARR
            Set fx = sl.TimeLine.MainSequence.AddEffect(Shape:=sp, _
                    effectId:=msoAnimEffectWipe, trigger:=trig)
            fx.Timing.Duration = 0.6
            SetWipeDirection fx, rec.motion
        Case KIND_TEXT
            Set fx = sl.TimeLine.MainSequence.AddEffect(Shape:=sp, _
                    effectId:=msoAnimEffectWipe, trigger:=trig)
            fx.Timing.Duration = 0.5
            SetWipeDirection fx, DIR_DOWN      ' text wipes top-to-bottom
        Case Else
            Set fx = sl.TimeLine.MainSequence.AddEffect(Shape:=sp, _
                    effectId:=msoAnimEffectFade, trigger:=trig)
            fx.Timing.Duration = 0.4
    End Select
End Sub

Private Sub SetWipeDirection(fx As Effect, motionDir As Long)
    ' NOTE: MsoAnimDirection values name where the wipe STARTS FROM, not the
    ' direction of motion.  i.e. msoAnimDirectionLeft = "wipe from left" =
    ' reveal begins at the left edge and sweeps rightward.  So to get motion
    ' in direction X we must pass the constant for the OPPOSITE edge.
    On Error Resume Next
    Dim d As MsoAnimDirection
    Select Case motionDir
        Case DIR_RIGHT: d = msoAnimDirectionLeft   ' from left  -> motion right
        Case DIR_LEFT:  d = msoAnimDirectionRight  ' from right -> motion left
        Case DIR_DOWN:  d = msoAnimDirectionUp     ' from top   -> motion down
        Case DIR_UP:    d = msoAnimDirectionDown   ' from bottom-> motion up
    End Select
    fx.EffectParameters.Direction = d
    On Error GoTo 0
End Sub
