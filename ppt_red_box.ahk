#Requires AutoHotkey v2.0
#SingleInstance Force
CoordMode "Mouse", "Screen"

appVersion := "1.0.66"
updateUrl := "https://raw.githubusercontent.com/hojung1231/ppt-red-box/main/ppt_red_box.ahk"
updateRequest := false
updateDeadline := 0
pendingUpdate := ""
updateStatus := ""
defaultFontName := "현대하모니 M"
defaultFontSize := 12
fontSettingsPath := A_ScriptDir "\ppt_red_box.ini"
outlineWeight := 1.5
outlineColors := [0x0000FF, 0xFF0000, 0x00FFFF, 0x00FF00, 0x000000] ; Red, blue, yellow, green, black in Office BGR
nativeFreeform := false
catiaRemoveFaceRepeat := false
catiaQuickTrimRepeat := false

if A_LineFile = A_ScriptFullPath
    InitializeApp()

#HotIf WinActive("ahk_exe POWERPNT.EXE")
^1:: {
    InsertPowerPointTable()
    KeyWait "1"
}
F3::ApplyDefaultFontToSelection()
^+x::InsertTextbox()
^+n:: {
    ApplyToSelectedShapes(ToggleShapeGlow)
    KeyWait "n"
}
^+SC029::CycleSelectedDashStyles()
$+WheelDown:: {
    if !ChangePowerPointTableSize("Rows", 1)
        SendEvent "+{WheelDown}"
}
$+WheelUp:: {
    if !ChangePowerPointTableSize("Rows", -1)
        SendEvent "+{WheelUp}"
}
^+WheelDown:: {
    if !ChangePowerPointTableSize("Columns", 1)
        AdjustSelectedOutlineWeight(-0.25)
}
^+WheelUp:: {
    if !ChangePowerPointTableSize("Columns", -1)
        AdjustSelectedOutlineWeight(0.25)
}
$^+q:: {
    if !CycleSelectedArrowheads(6) ; msoArrowheadOval
        Send "^+q"
    KeyWait "q"
}
$^+a:: {
    if !CycleSelectedArrowheads()
        Send "^+a"
}
$^+c:: {
    if !CycleSelectedColors()
        Send "^+c"
}
MButton:: {
    MouseGetPos &startX, &startY
    TrackNativeFreeform()
    try {
        canvas := GetSlideCanvas()
        if !canvas
            return

        ; The outline on screen is temporary. Only the final shape goes on the slide.
        borders := []
        dragged := false
        lastX := startX, lastY := startY
        try {
            while GetKeyState("MButton", "P") {
                if !WinActive("ahk_id " canvas.hwnd)
                    return
                MouseGetPos &x, &y
                if x != lastX || y != lastY {
                    if dragged || IsDrag(startX, startY, x, y) {
                        dragged := true
                        if !borders.Length
                            borders := CreatePreviewBorders()
                        ShowRectanglePreview(borders, startX, startY, x, y,
                            Max(1, Round(outlineWeight / canvas.scaleX)))
                    }
                    lastX := x, lastY := y
                }
                Sleep 10
            }
            MouseGetPos &x, &y
        } finally {
            for border in borders
                border.Destroy()
        }

        current := GetSlideCanvas()
        if !current || !IsSameSlide(canvas, current)
            return
        if dragged || IsDrag(startX, startY, x, y) {
            if nativeFreeform {
                Send "{Esc}"
                StopTrackingFreeform()
            }
            DrawRectangle(current, startX, startY, x, y)
        } else if !nativeFreeform {
            StartNativeFreeform(current)
        }
    } catch {
        ShowHint("PowerPoint 편집 화면에서 다시 시도하세요.")
    }
}

#HotIf nativeFreeform && WinActive("ahk_id " nativeFreeform.hwnd)
~Esc Up::StopTrackingFreeform()
#HotIf

#HotIf (catiaRemoveFaceRepeat || catiaQuickTrimRepeat) && WinActive("ahk_exe CNEXT.exe")
~Esc:: {
    StopCatiaRemoveFaceRepeat()
    StopCatiaQuickTrimRepeat()
}
#HotIf catiaQuickTrimRepeat && WinActive("ahk_exe CNEXT.exe")
~LButton Up::QueueCatiaQuickTrim()
#HotIf WinActive("ahk_exe CNEXT.exe")
^Numpad5::
^NumpadClear:: {
    CycleCatiaUpperIsometricView()
    KeyWait "sc04C"
}
^Numpad8::
^NumpadUp::SetCatiaDirectionalView(0, -1)
^Numpad2::
^NumpadDown::SetCatiaDirectionalView(0, 1)
^Numpad4::
^NumpadLeft::SetCatiaDirectionalView(1, 0)
^Numpad6::
^NumpadRight::SetCatiaDirectionalView(-1, 0)
^Numpad7::
^NumpadHome::SetCatiaDirectionalView(1, -1)
^Numpad9::
^NumpadPgUp::SetCatiaDirectionalView(-1, -1)
^Numpad1::
^NumpadEnd::SetCatiaDirectionalView(1, 1)
^Numpad3::
^NumpadPgDn::SetCatiaDirectionalView(-1, 1)
~m:: {
    if A_PriorHotkey = "~m" && A_TimeSincePriorHotkey < 350 {
        KeyWait "m"
        StartCatiaMeasureBetween()
    }
    KeyWait "m"
}
~q:: {
    if A_PriorHotkey = "~q" && A_TimeSincePriorHotkey < 350 {
        KeyWait "q"
        StartCatiaQuickTrim()
    }
    KeyWait "q"
}
~j:: {
    if A_PriorHotkey = "~j" && A_TimeSincePriorHotkey < 350 {
        KeyWait "j"
        StartCatiaGsdCommand("Join")
    }
    KeyWait "j"
}
~x:: {
    if A_PriorHotkey = "~x" && A_TimeSincePriorHotkey < 350 {
        KeyWait "x"
        StartCatiaGsdCommand("Extrapolate")
    }
    KeyWait "x"
}
~s:: {
    if A_PriorHotkey = "~s" && A_TimeSincePriorHotkey < 350 {
        KeyWait "s"
        StartCatiaSplit()
    }
    KeyWait "s"
}
~v:: {
    if A_PriorHotkey = "~v" && A_TimeSincePriorHotkey < 350 {
        KeyWait "v"
        ToggleCatiaView()
    }
    KeyWait "v"
}
~l:: {
    if A_PriorHotkey = "~l" && A_TimeSincePriorHotkey < 350 {
        KeyWait "l"
        StartCatiaLine()
    }
    KeyWait "l"
}
^e:: {
    KeyWait "e"
    StartCatiaGsdCommand("Extrude")
}
~e:: {
    if A_PriorHotkey = "~e" && A_TimeSincePriorHotkey < 350 {
        KeyWait "e"
        StartCatiaGsdCommand("Multiple Extract")
    }
    KeyWait "e"
}
~c:: {
    if A_PriorHotkey = "~c" && A_TimeSincePriorHotkey < 350 {
        KeyWait "c"
        ApplyCatiaSketchCoincidence()
    }
    KeyWait "c"
}
~d:: {
    if A_PriorHotkey = "~d" && A_TimeSincePriorHotkey < 350 {
        KeyWait "d"
        StartCatiaSketchPoint()
    }
    KeyWait "d"
}
^t:: {
    KeyWait "t"
    StartCatiaThickness()
}
~t:: {
    if A_PriorHotkey = "~t" && A_TimeSincePriorHotkey < 350 {
        KeyWait "t"
        StartCatiaThickSurface()
    }
    KeyWait "t"
}
~r:: {
    if A_PriorHotkey = "~r" && A_TimeSincePriorHotkey < 350 {
        KeyWait "r"
        StartCatiaRemoveFaceRepeat()
    }
    KeyWait "r"
}
~i:: {
    if A_PriorHotkey = "~i" && A_TimeSincePriorHotkey < 350 {
        KeyWait "i"
        StartCatiaIntersectOrIsolate()
    }
    KeyWait "i"
}
~p:: {
    if A_PriorHotkey = "~p" && A_TimeSincePriorHotkey < 350
        StartCatiaPlaneOrProjection()
    KeyWait "p"
}
~h:: {
    if A_PriorHotkey = "~h" && A_TimeSincePriorHotkey < 350 {
        if !ToggleCatiaConstruction()
            ToggleCatiaContainerAndOpenSubtree()
    }
    KeyWait "h"
}
^f:: {
    KeyWait "f"
    DefineCatiaSelectedContainer()
}
$f:: {
    if A_PriorHotkey = "$f" && A_TimeSincePriorHotkey < 350 {
        KeyWait "f"
        StartCatiaGsdCommand("Fill")
    } else if !ToggleCatiaSketchFix() {
        Send "f"
    }
    KeyWait "f"
}
^r:: {
    OpenCatiaReorderChildren()
    KeyWait "r"
}
^q:: {
    CreateCatiaContainer("Body")
    KeyWait "q"
}
^w:: {
    CreateCatiaContainer("Geometrical Set")
    KeyWait "w"
}
^3:: {
    ToggleCatiaVisibility("Wireframe3D")
    KeyWait "3"
}
^2:: {
    ToggleCatiaVisibility("Sketch")
    KeyWait "2"
}
^1:: {
    ToggleCatiaVisibility("Body")
    KeyWait "1"
}
#HotIf

FindCatiaContainer(original) {
    container := 0
    candidate := original
    Loop 30 {
        ; Parent paths include collections such as HybridShapes, which
        ; CATIA cannot add to Selection. Inspect containers directly.
        try {
            shapes := candidate.Shapes
            sketches := candidate.Sketches
            container := candidate ; Body
        }
        if !container {
            try {
                shapes := candidate.HybridShapes
                sketches := candidate.HybridSketches
                container := candidate ; HybridBody / Geometrical Set
            }
        }
        if container
            break
        candidate := candidate.Parent
    }
    if !container
        throw Error("상위 Body 또는 Geometrical Set을 찾을 수 없습니다.")
    return container
}

FindCatiaOwningPart(container) {
    candidate := container
    Loop 30 {
        try {
            origin := candidate.OriginElements
            bodies := candidate.Bodies
            return candidate
        }
        candidate := candidate.Parent
    }
    throw Error("선택한 개체가 속한 Part를 찾을 수 없습니다.")
}

DefineCatiaSelectedContainer() {
    try {
        catia := ComObjActive("CATIA.Application")
        document := catia.ActiveDocument
        selection := document.Selection
        if selection.Count2 != 1 {
            ShowHint("화면 또는 트리에서 개체 하나를 선택하세요.")
            return
        }
        container := FindCatiaContainer(selection.Item2(1).Value)
        part := FindCatiaOwningPart(container)
        part.InWorkObject := container
        ShowHint(container.Name " · Define In Work Object")
    } catch {
        ShowHint("선택한 개체의 Body 또는 Geometrical Set을 작업 대상으로 지정할 수 없습니다.")
    }
}

StartCatiaRemoveFaceRepeat() {
    global catiaRemoveFaceRepeat
    StopCatiaRemoveFaceRepeat()
    try {
        catia := ComObjActive("CATIA.Application")
        document := catia.ActiveDocument
        part := document.Part
        catiaRemoveFaceRepeat := {document: document, waiting: true,
            deadline: A_TickCount + 5000}
        catia.StartCommand("Remove Face")
        SetTimer WatchCatiaRemoveFace, 150
    } catch {
        StopCatiaRemoveFaceRepeat()
        ShowHint("CATIA CATPart에서 Remove Face를 실행할 수 없습니다.")
    }
}

WatchCatiaRemoveFace() {
    global catiaRemoveFaceRepeat
    if !catiaRemoveFaceRepeat
        return
    dialog := WinExist("Remove Face Definition ahk_exe CNEXT.exe")
        || WinExist("Remove Face/Edge Definition ahk_exe CNEXT.exe")
    if dialog {
        catiaRemoveFaceRepeat.waiting := false
        return
    }
    if catiaRemoveFaceRepeat.waiting {
        if A_TickCount >= catiaRemoveFaceRepeat.deadline {
            StopCatiaRemoveFaceRepeat()
            ShowHint("Remove Face 창을 확인할 수 없어 연속 실행을 종료했습니다.")
        }
        return
    }
    if !WinActive("ahk_exe CNEXT.exe") {
        StopCatiaRemoveFaceRepeat()
        return
    }
    try {
        catia := ComObjActive("CATIA.Application")
        if ComObjValue(catia.ActiveDocument) != ComObjValue(catiaRemoveFaceRepeat.document) {
            StopCatiaRemoveFaceRepeat()
            return
        }
        catiaRemoveFaceRepeat.waiting := true
        catiaRemoveFaceRepeat.deadline := A_TickCount + 5000
        catia.StartCommand("Remove Face")
    } catch {
        StopCatiaRemoveFaceRepeat()
    }
}

StopCatiaRemoveFaceRepeat(*) {
    global catiaRemoveFaceRepeat
    SetTimer WatchCatiaRemoveFace, 0
    catiaRemoveFaceRepeat := false
}

ToggleCatiaConstruction() {
    try {
        catia := ComObjActive("CATIA.Application")
        selection := catia.ActiveDocument.Selection
        if !selection.Count2
            return false
        Loop selection.Count2 {
            geometry := selection.Item2(A_Index).Value
            ; Identify sketch geometry without changing individual states.
            construction := geometry.Construction
        }
    } catch {
        return false
    }
    SetTimer(RunCatiaConstructionCommand.Bind(catia), -100)
    return true
}

RunCatiaConstructionCommand(catia) {
    KeyWait "h"
    try {
        ; Use the Sketch Tools command so CATIA handles mixed selections
        ; and the relationship between curves and their points.
        catia.StartCommand("Construction/Standard Element")
    } catch {
        ShowHint("スケ치의 Construction/Standard Element를 실행할 수 없습니다.")
    }
}

ToggleCatiaContainerAndOpenSubtree() {
    selection := 0
    original := 0
    container := 0
    try {
        catia := ComObjActive("CATIA.Application")
        selection := catia.ActiveDocument.Selection
        if selection.Count2 != 1 {
            ShowHint("화면의 솔리드·면 또는 트리의 Body·Geometrical Set 하나를 선택하세요.")
            return
        }
        original := selection.Item2(1).Value
        container := FindCatiaContainer(original)
        hideTarget := container
        try {
            ; Bodies under an OGS can ignore their own No Show property.
            ; Their terminal solid feature controls the displayed result.
            orderedSketches := container.Parent.Parent.OrderedSketches
            shapes := container.Shapes
            if shapes.Count
                hideTarget := shapes.Item(shapes.Count)
        }
        selection.Clear()
        selection.Add(hideTarget)
        state := Buffer(4, 0)
        selection.VisProperties.GetShow(ComValue(0x4003, state.Ptr))
        nextState := NumGet(state, 0, "Int") = 1 ? 0 : 1
        selection.VisProperties.SetShow(nextState)
        selection.Clear()
        selection.Add(container)
        ; CATIA dispatches contextual commands asynchronously. Keep the
        ; container selected and invoke the command after the hotkey ends.
        SetTimer(OpenCatiaSubtree.Bind(catia, selection, container), -100)
    } catch as error {
        if selection && original {
            try {
                selection.Clear()
                selection.Add(original)
            }
        }
        ShowHint(error.Message)
    }
}

OpenCatiaSubtree(catia, selection, container) {
    KeyWait "h"
    try {
        selection.Clear()
        selection.Add(container)
        catia.StartCommand("OpenNewSubTree")
        ; Leave this selection intact for CATIA's command event loop.
    } catch {
        ShowHint("서브트리를 열 수 없습니다. 선택된 컨테이너에서 Open Sub-Tree를 확인하세요.")
    }
}

StartCatiaMeasureBetween() {
    try {
        catia := ComObjActive("CATIA.Application")
        catia.StartCommand("Measure Between")
    } catch {
        ShowHint("CATIA에서 Measure Between을 실행할 수 없습니다.")
    }
}

StartCatiaThickSurface() {
    try {
        catia := ComObjActive("CATIA.Application")
        catia.StartCommand("Thick Surface")
    } catch {
        ShowHint("CATIA에서 Thick Surface를 실행할 수 없습니다.")
    }
}

StartCatiaThickness() {
    try {
        catia := ComObjActive("CATIA.Application")
        catia.StartCommand("Thickness")
    } catch {
        ShowHint("CATIA에서 Thickness를 실행할 수 없습니다.")
    }
}

StartCatiaIntersectOrIsolate() {
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() = "CS0WKS"
            catia.StartCommand("Intersect 3D Elements")
        else
            IsolateCatiaSelection()
    } catch {
        ShowHint("Intersect 3D Elements 또는 Isolate를 실행할 수 없습니다.")
    }
}

IsolateCatiaSelection() {
    try {
        catia := ComObjActive("CATIA.Application")
        if !catia.ActiveDocument.Selection.Count2 {
            ShowHint("Isolate할 개체를 먼저 선택하세요.")
            return
        }
        catia.StartCommand("Isolate")
    } catch {
        ShowHint("선택한 개체에서 Isolate를 실행할 수 없습니다.")
    }
}

StartCatiaGsdCommand(command) {
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() = "CS0WKS" {
            ShowHint("스케치에서 나와 GSD에서 사용하세요.")
            return
        }
        catia.StartCommand(command)
    } catch {
        ShowHint("GSD 명령을 실행할 수 없습니다: " command)
    }
}

StartCatiaQuickTrim() {
    global catiaQuickTrimRepeat
    StopCatiaQuickTrimRepeat()
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() != "CS0WKS" {
            ShowHint("스케치 편집 중에 QQ를 사용하세요.")
            return
        }
        MouseGetPos(, , &window, &viewControl, 2)
        catiaQuickTrimRepeat := {document: catia.ActiveDocument,
            window: window, viewControl: viewControl}
        catia.StartCommand("Quick Trim")
    } catch {
        StopCatiaQuickTrimRepeat()
        ShowHint("스케치에서 Quick Trim을 실행할 수 없습니다.")
    }
}

QueueCatiaQuickTrim() {
    if !catiaQuickTrimRepeat
        return
    MouseGetPos(, , &window, &viewControl, 2)
    if window != catiaQuickTrimRepeat.window
        || viewControl != catiaQuickTrimRepeat.viewControl {
        StopCatiaQuickTrimRepeat()
        return
    }
    SetTimer RearmCatiaQuickTrim, -150
}

RearmCatiaQuickTrim() {
    if !catiaQuickTrimRepeat
        return
    if !WinActive("ahk_id " catiaQuickTrimRepeat.window) {
        StopCatiaQuickTrimRepeat()
        return
    }
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() != "CS0WKS"
            || ComObjValue(catia.ActiveDocument) != ComObjValue(catiaQuickTrimRepeat.document) {
            StopCatiaQuickTrimRepeat()
            return
        }
        catia.StartCommand("Quick Trim")
    } catch {
        StopCatiaQuickTrimRepeat()
    }
}

StopCatiaQuickTrimRepeat(*) {
    global catiaQuickTrimRepeat
    SetTimer RearmCatiaQuickTrim, 0
    catiaQuickTrimRepeat := false
}

FindCatiaSelectedSketch(part, selection) {
    try {
        sketch := part.InWorkObject
        elements := sketch.GeometricElements
        constraints := sketch.Constraints
        return sketch
    }
    ; InWorkObject may remain the containing Geometrical Set during
    ; sketch editing. Follow selected geometry back to its Sketch.
    Loop selection.Count2 {
        candidate := selection.Item2(A_Index).Value
        Loop 20 {
            try {
                elements := candidate.GeometricElements
                constraints := candidate.Constraints
                return candidate
            }
            try candidate := candidate.Parent
            catch
                break
        }
    }
    throw Error("선택 개체가 속한 스케치를 찾을 수 없습니다.")
}

ApplyCatiaSketchCoincidence() {
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() != "CS0WKS" {
            ShowHint("스케치 편집 중에 CC를 사용하세요.")
            return
        }
        document := catia.ActiveDocument
        selection := document.Selection
        if selection.Count2 != 2 {
            ShowHint("일치시킬 점·선 등 두 개체를 선택하고 CC를 누르세요.")
            return
        }
        sketch := FindCatiaSelectedSketch(document.Part, selection)
        constraints := sketch.Constraints
        first := selection.Item2(1).Reference
        second := selection.Item2(2).Reference
        Loop constraints.Count {
            constraint := constraints.Item(A_Index)
            if constraint.Type != 2
                continue
            a := constraint.GetConstraintElement(1).DisplayName
            b := constraint.GetConstraintElement(2).DisplayName
            if (a = first.DisplayName && b = second.DisplayName)
                || (a = second.DisplayName && b = first.DisplayName) {
                ShowHint("이미 Coincidence 구속이 적용돼 있습니다.")
                return
            }
        }
        constraints.AddBiEltCst(2, first, second) ; catCstTypeOn
        ShowHint("Coincidence 적용")
    } catch {
        ShowHint("선택한 두 개체에 Coincidence 구속을 적용할 수 없습니다.")
    }
}

StartCatiaSplit() {
    try {
        catia := ComObjActive("CATIA.Application")
        workbench := catia.GetWorkbenchId()
        if workbench = "CS0WKS" {
            ShowHint("스케치에서 나와 Part Design 또는 GSD에서 SS를 사용하세요.")
            return
        }
        if workbench = "PrtCfg" {
            ; Both workbenches register "Split". Select the Part Design
            ; menu explicitly to avoid dispatching the GSD surface command.
            hwnd := WinActive("ahk_exe CNEXT.exe")
            MenuSelect("ahk_id " hwnd, , "Insert", "Surface-Based Features", "Split...")
        } else {
            catia.StartCommand("Split")
        }
    } catch {
        ShowHint("Split을 실행할 수 없습니다. Insert → Surface-Based Features → Split 메뉴를 확인하세요.")
    }
}

CycleCatiaUpperIsometricView() {
    static index := 1
    directions := [[-1, -1], [1, -1], [1, 1], [-1, 1]]
    direction := directions[index]
    if SetCatiaDirectionalView(direction[2], -1, direction[1])
        index := Mod(index, directions.Length) + 1
}

SetCatiaDirectionalView(sightY, sightZ, sightX := 0) {
    try {
        catia := ComObjActive("CATIA.Application")
        script := "
        (
        Function SetDirection(sightY, sightZ, sightX)
            Dim viewer, view, origin(2), sight(2), up(2)
            Dim target(2), distance, magnitude, oldMagnitude, i
            Set viewer = CATIA.ActiveWindow.ActiveViewer
            Set view = viewer.Viewpoint3D
            view.GetOrigin origin
            view.GetSightDirection sight
            distance = view.FocusDistance
            oldMagnitude = Sqr(sight(0)^2 + sight(1)^2 + sight(2)^2)
            If oldMagnitude = 0 Then Err.Raise 5, , "Invalid sight direction"
            For i = 0 To 2
                target(i) = origin(i) + distance * sight(i) / oldMagnitude
            Next
            magnitude = Sqr(sightX^2 + sightY^2 + sightZ^2)
            sight(0) = sightX / magnitude
            sight(1) = sightY / magnitude
            sight(2) = sightZ / magnitude
            If Abs(sight(2)) > 0.999 Then
                up(0) = -1
                up(1) = 0
                up(2) = 0
            Else
                up(0) = -sight(2) * sight(0)
                up(1) = -sight(2) * sight(1)
                up(2) = 1 - sight(2)^2
                magnitude = Sqr(up(0)^2 + up(1)^2 + up(2)^2)
                up(0) = up(0) / magnitude
                up(1) = up(1) / magnitude
                up(2) = up(2) / magnitude
            End If
            For i = 0 To 2
                origin(i) = target(i) - distance * sight(i)
            Next
            view.PutOrigin origin
            view.PutSightDirection sight
            view.PutUpDirection up
            viewer.Update
            SetDirection = True
        End Function
        )"
        parameters := ComObjArray(12, 3)
        parameters[0] := sightY
        parameters[1] := sightZ
        parameters[2] := sightX
        return catia.SystemService.Evaluate(script, 0, "SetDirection", parameters)
    } catch {
        ShowHint("CATIA 방향 뷰를 설정할 수 없습니다.")
        return false
    }
}

ToggleCatiaView() {
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() = "CS0WKS" {
            catia.StartCommand("Normal View")
            return
        }
        ; CATIA's vector output arguments are handled inside VBScript.
        script := "
        (
        Function ReverseView()
            Dim viewer, view, origin(2), sight(2), up(2), distance, magnitude, i
            Set viewer = CATIA.ActiveWindow.ActiveViewer
            Set view = viewer.Viewpoint3D
            view.GetOrigin origin
            view.GetSightDirection sight
            view.GetUpDirection up
            distance = view.FocusDistance
            magnitude = Sqr(sight(0)^2 + sight(1)^2 + sight(2)^2)
            If magnitude = 0 Then Err.Raise 5, , "Invalid sight direction"
            For i = 0 To 2
                origin(i) = origin(i) + 2 * distance * sight(i) / magnitude
                sight(i) = -sight(i)
            Next
            view.PutOrigin origin
            view.PutSightDirection sight
            view.PutUpDirection up
            viewer.Update
            ReverseView = True
        End Function
        )"
        catia.SystemService.Evaluate(script, 0, "ReverseView", ComObjArray(12, 0))
    } catch {
        ShowHint("CATIA에서 뷰 방향을 전환할 수 없습니다.")
    }
}

StartCatiaLine() {
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() = "CS0WKS"
            catia.StartCommand("Line")
        else
            catia.StartCommand("Line...")
    } catch {
        ShowHint("CATIA에서 Line을 실행할 수 없습니다.")
    }
}

StartCatiaSketchPoint() {
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() != "CS0WKS" {
            ShowHint("스케치 편집 중에 DD를 사용하세요.")
            return
        }
        catia.StartCommand("Point")
    } catch {
        ShowHint("스케치에서 Point by Clicking을 실행할 수 없습니다.")
    }
}

StartCatiaPlaneOrProjection() {
    try {
        catia := ComObjActive("CATIA.Application")
        if catia.GetWorkbenchId() = "CS0WKS"
            catia.StartCommand("Project 3D Elements")
        else
            catia.StartCommand("Plane")
    } catch {
        ShowHint("CATIA에서 Plane 또는 Project 3D Elements를 실행할 수 없습니다.")
    }
}

ToggleCatiaSketchFix() {
    try {
        catia := ComObjActive("CATIA.Application")
        document := catia.ActiveDocument
        part := document.Part
        selection := document.Selection
        if selection.Count2 != 1
            return false
        geometry := selection.Item2(1).Value
        sketch := part.InWorkObject
        ; These properties exist on a Sketch, not on a solid Body.
        elements := sketch.GeometricElements
        constraints := sketch.Constraints
        reference := part.CreateReferenceFromObject(geometry)
        belongsToSketch := false
        Loop elements.Count {
            element := elements.Item(A_Index)
            elementReference := part.CreateReferenceFromObject(element)
            if elementReference.DisplayName = reference.DisplayName {
                belongsToSketch := true
                break
            }
        }
        if !belongsToSketch
            return false
    } catch {
        return false
    }
    try {
        fixes := []
        Loop constraints.Count {
            constraint := constraints.Item(A_Index)
            if constraint.Type = 0 { ; catCstTypeReference (Fix)
                constrained := constraint.GetConstraintElement(1)
                if constrained.DisplayName = reference.DisplayName
                    fixes.Push(constraint)
            }
        }
        if fixes.Length {
            try {
                selection.Clear()
                for constraint in fixes
                    selection.Add(constraint)
                selection.Delete()
            } finally {
                selection.Clear()
                selection.Add(geometry)
            }
            ShowHint("Fix 해제")
        } else {
            constraints.AddMonoEltCst(0, reference)
            ShowHint("Fix 적용")
        }
    } catch {
        ShowHint("선택한 스케치 개체의 Fix 구속을 변경할 수 없습니다.")
    }
    return true
}

OpenCatiaReorderChildren() {
    try {
        catia := ComObjActive("CATIA.Application")
        document := catia.ActiveDocument
        part := document.Part
        selection := document.Selection
        selection.Clear()
        selection.Add(part)
        ; Keep the root selected while CATIA opens its contextual command.
        catia.StartCommand("Reorder Children")
    } catch {
        ShowHint("CATIA의 CATPart를 열고 다시 시도하세요.")
    }
}

CreateCatiaContainer(kind) {
    try {
        catia := ComObjActive("CATIA.Application")
        part := catia.ActiveDocument.Part
        if kind = "Body"
            created := part.Bodies.Add()
        else
            created := part.HybridBodies.Add()
        part.InWorkObject := created
    } catch {
        ShowHint("CATIA V5의 CATPart에서 " kind "를 만들 수 없습니다.")
    }
}

ToggleCatiaVisibility(kind) {
    try {
        catia := ComObjActive("CATIA.Application")
        document := catia.ActiveDocument
        part := document.Part

        selection := document.Selection
        saved := []
        Loop selection.Count2
            saved.Push(selection.Item2(A_Index).Value)
        try {
            items := []
            if kind = "Sketch" || kind = "Wireframe3D" {
                selection.Clear()
                query := kind = "Wireframe3D"
                    ? "(CATPrtSearch.Point + CATPrtSearch.Line + CATPrtSearch.Curve + CATGmoSearch.GSMIntersect + CATGmoSearch.GSMProject),all"
                    : "CATPrtSearch.Sketch,all"
                selection.Search(query)
                Loop selection.Count2
                    items.Push(selection.Item2(A_Index).Value)
            } else {
                bodies := part.Bodies
                Loop bodies.Count
                    items.Push(bodies.Item(A_Index))
            }
            if !items.Length
                return
            target := 0 ; Show all when every item is hidden.
            for item in items {
                selection.Clear()
                selection.Add(item)
                state := Buffer(4, 0)
                selection.VisProperties.GetShow(ComValue(0x4003, state.Ptr))
                if NumGet(state, 0, "Int") = 0
                    target := 1 ; Hide all when any item is shown.
            }
            selection.Clear()
            for item in items
                selection.Add(item)
            selection.VisProperties.SetShow(target)
        } finally {
            selection.Clear()
            for item in saved
                try selection.Add(item)
        }
    } catch {
        ShowHint("CATIA V5의 CATPart를 열고 다시 시도하세요.")
    }
}

IsDrag(x1, y1, x2, y2) {
    return Abs(x2 - x1) > 5 || Abs(y2 - y1) > 5
}

CreatePreviewBorders() {
    borders := []
    Loop 4
        borders.Push(CreatePreviewWindow())
    return borders
}

CreatePreviewWindow() {
    preview := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x08080020")
    preview.BackColor := "FF0000"
    DllCall("SetLayeredWindowAttributes", "Ptr", preview.Hwnd, "UInt", 0, "UChar", 255, "UInt", 2)
    return preview
}

ShowRectanglePreview(borders, x1, y1, x2, y2, thickness) {
    left := Min(x1, x2), top := Min(y1, y2)
    width := Max(thickness, Abs(x2 - x1))
    height := Max(thickness, Abs(y2 - y1))
    borders[1].Show(Format("NA x{} y{} w{} h{}", left, top, width, thickness))
    borders[2].Show(Format("NA x{} y{} w{} h{}", left, top + height - thickness, width, thickness))
    borders[3].Show(Format("NA x{} y{} w{} h{}", left, top, thickness, height))
    borders[4].Show(Format("NA x{} y{} w{} h{}", left + width - thickness, top, thickness, height))
}

GetSlideCanvas() {
    hwnd := WinActive("ahk_exe POWERPNT.EXE")
    if !hwnd
        return false
    ppt := ComObjActive("PowerPoint.Application")
    win := ppt.ActiveWindow
    slide := win.View.Slide
    originX := win.PointsToScreenPixelsX(0), originY := win.PointsToScreenPixelsY(0)
    return {hwnd: hwnd, slide: slide, slideId: slide.SlideID,
        presentation: win.Presentation, presentationName: win.Presentation.FullName,
        originX: originX, originY: originY,
        scaleX: 100 / (win.PointsToScreenPixelsX(100) - originX),
        scaleY: 100 / (win.PointsToScreenPixelsY(100) - originY)}
}

IsSameSlide(first, second) {
    return first.hwnd = second.hwnd && first.presentationName = second.presentationName
        && first.slideId = second.slideId
}

ToSlidePoint(canvas, x, y) {
    return {x: (x - canvas.originX) * canvas.scaleX, y: (y - canvas.originY) * canvas.scaleY}
}

StylePowerPointTableCell(cell, row, column) {
    fill := cell.Shape.Fill
    if row = 1 || column = 1 {
        fill.Solid()
        fill.Visible := -1
        fill.ForeColor.RGB := 0xF2F2F2
        fill.Transparency := 0
    } else {
        fill.Visible := 0
    }
    frame := cell.Shape.TextFrame
    frame.VerticalAnchor := 3 ; msoAnchorMiddle
    text := frame.TextRange
    text.Font.Color.RGB := 0x000000
    ApplyDefaultFont(text.Font)
    text.ParagraphFormat.Alignment := 2 ; ppAlignCenter
    Loop 4 {
        border := cell.Borders.Item(A_Index)
        border.Visible := -1
        border.ForeColor.RGB := 0x000000
        border.Transparency := 0
        border.Weight := 0.5
    }
}

ChangePowerPointTableSize(dimension, delta) {
    try {
        ppt := ComObjActive("PowerPoint.Application")
        selected := ppt.ActiveWindow.Selection.ShapeRange
        if selected.Count != 1 || selected.Item(1).HasTable != -1
            return false
        tableShape := selected.Item(1)
        table := tableShape.Table
    } catch {
        return false
    }
    try {
        entries := dimension = "Rows" ? table.Rows : table.Columns
        if delta < 0 {
            if entries.Count > 1
                entries.Item(entries.Count).Delete()
            return true
        }
        widths := []
        heights := []
        if dimension = "Columns" {
            left := tableShape.Left
            top := tableShape.Top
            Loop table.Columns.Count
                widths.Push(table.Columns.Item(A_Index).Width)
            Loop table.Rows.Count
                heights.Push(table.Rows.Item(A_Index).Height)
        }
        ; Native Add inherits the adjacent row/column's cell formatting.
        entries.Add()
        if dimension = "Columns" {
            ; Preserve all old dimensions while appending a full-width column.
            for index, width in widths
                table.Columns.Item(index).Width := width
            table.Columns.Item(table.Columns.Count).Width := widths[widths.Length]
            for index, height in heights
                table.Rows.Item(index).Height := height
            tableShape.Left := left
            tableShape.Top := top
        }
    } catch {
        ShowHint("표의 행·열을 변경할 수 없습니다.")
    }
    return true
}

InsertPowerPointTable() {
    TrackNativeFreeform()
    try {
        canvas := GetSlideCanvas()
        if !canvas
            return
        pointsPerCentimeter := 72 / 2.54
        width := 15 * pointsPerCentimeter
        height := 4 * pointsPerCentimeter
        page := canvas.presentation.PageSetup
        left := Max(0, (page.SlideWidth - width) / 2)
        top := Max(0, (page.SlideHeight - height) / 2)
        shape := canvas.slide.Shapes.AddTable(5, 3, left, top, width, height)
        try {
            table := shape.Table
            Loop table.Rows.Count {
                row := A_Index
                Loop table.Columns.Count {
                    column := A_Index
                    StylePowerPointTableCell(table.Cell(row, column), row, column)
                }
            }
        } catch as error {
            shape.Delete()
            throw error
        }
        shape.Select()
        try shape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Select()
    } catch {
        ShowHint("표를 만들 수 없습니다. PowerPoint 편집 화면을 확인하세요.")
    }
}

InsertTextbox() {
    TrackNativeFreeform()
    try {
        canvas := GetSlideCanvas()
        if !canvas
            return

        MouseGetPos &x, &y
        point := ToSlidePoint(canvas, x, y)
        page := canvas.presentation.PageSetup
        shape := CreateTextbox(canvas.slide.Shapes, point, page.SlideWidth, page.SlideHeight)
        shape.Select()
        try shape.TextFrame.TextRange.Select()
    } catch {
        ShowHint("텍스트 상자를 만들 수 없습니다. PowerPoint 편집 화면을 확인하세요.")
    }
}

CreateTextbox(shapes, point, slideWidth, slideHeight) {
    left := Max(0, Min(point.x, slideWidth - 1))
    top := Max(0, Min(point.y, slideHeight - 1))
    ; AddTextbox requires initial dimensions; AutoSize replaces this seed
    ; size as text is entered, like a click-created native text box.
    shape := shapes.AddTextbox(1, left, top, 1, 1)
    try {
        shape.Fill.Visible := 0
        shape.Line.Visible := 0
        frame := shape.TextFrame
        frame.WordWrap := 0
        frame.AutoSize := 1 ; ppAutoSizeShapeToFitText
        frame.TextRange.ParagraphFormat.Alignment := 2 ; ppAlignCenter
        font := frame.TextRange.Font
        ApplyDefaultFont(font)
        font.Color.RGB := 0x000000
        ApplyWhiteGlow(shape.TextFrame2.TextRange.Font.Glow, 10, 0.10)
    } catch as error {
        shape.Delete()
        throw error
    }
    return shape
}

ApplyDefaultFontToSelection() {
    TrackNativeFreeform()
    try {
        ppt := ComObjActive("PowerPoint.Application")
        selection := ppt.ActiveWindow.Selection
        switch selection.Type {
            case 3: ; ppSelectionText
                ApplyDefaultFont(selection.TextRange.Font)
            case 2: ; ppSelectionShapes
                selected := selection.ShapeRange
                Loop selected.Count {
                    shape := selected.Item(A_Index)
                    try {
                        if shape.HasTextFrame = -1
                            ApplyDefaultFont(shape.TextFrame.TextRange.Font)
                    }
                }
        }
    }
}

ApplyDefaultFont(font) {
    font.Name := defaultFontName
    font.NameFarEast := defaultFontName
    font.Size := defaultFontSize
}

LoadFontSettings() {
    global defaultFontName, defaultFontSize
    try {
        name := IniRead(fontSettingsPath, "Text", "Font", defaultFontName)
        size := IniRead(fontSettingsPath, "Text", "Size", defaultFontSize)
        if name = "현대하모니 M" || name = "NanumBarunGothicOTF"
            defaultFontName := name
        if IsNumber(size) && size > 0
            defaultFontSize := Number(size)
    }
}

SaveFontSettings(fontChoice, sizeChoice, status, *) {
    global defaultFontName, defaultFontSize
    size := Trim(sizeChoice.Text)
    if !IsNumber(size) || size <= 0 {
        status.Text := "글씨 크기를 양수로 입력하세요."
        return
    }
    name := fontChoice.Value = 1 ? "현대하모니 M" : "NanumBarunGothicOTF"
    try {
        IniWrite("Font=" name "`nSize=" size, fontSettingsPath, "Text")
        defaultFontName := name
        defaultFontSize := Number(size)
        status.Text := "저장됨"
    } catch {
        status.Text := "설정 파일을 저장할 수 없습니다."
    }
}

DrawRectangle(canvas, x1, y1, x2, y2) {
    start := ToSlidePoint(canvas, x1, y1)
    finish := ToSlidePoint(canvas, x2, y2)
    shape := canvas.slide.Shapes.AddShape(1, Min(start.x, finish.x), Min(start.y, finish.y),
        Abs(finish.x - start.x), Abs(finish.y - start.y))
    StyleOutline(shape)
    shape.Select()
}

StartNativeFreeform(canvas) {
    global nativeFreeform
    ppt := ComObjActive("PowerPoint.Application")
    shapes := canvas.slide.Shapes
    originalCount := shapes.Count
    originalIds := Map()
    Loop originalCount
        originalIds[shapes.Item(A_Index).Id] := true

    ActivateFreeformTool(ppt.CommandBars)
    nativeFreeform := {hwnd: canvas.hwnd, slideId: canvas.slideId,
        presentationName: canvas.presentationName, originalIds: originalIds}
    SetTimer TrackNativeFreeform, 100
}

ActivateFreeformTool(commandBars) {
    try commandBars.ExecuteMso("ShapeFreeform")
    catch {
        ; PowerPoint versions that lack the ribbon command can expose its legacy control.
        control := commandBars.FindControl(, 200)
        control.Execute()
    }
}

TrackNativeFreeform() {
    if !nativeFreeform
        return
    if !WinActive("ahk_id " nativeFreeform.hwnd) {
        StopTrackingFreeform()
        return
    }
    if GetKeyState("LButton", "P")
        return

    ; PowerPoint can temporarily reject automation while its drawing tool is handling input.
    try {
        ppt := ComObjActive("PowerPoint.Application")
        win := ppt.ActiveWindow
        slide := win.View.Slide
        if slide.SlideID != nativeFreeform.slideId || win.Presentation.FullName != nativeFreeform.presentationName {
            StopTrackingFreeform()
            return
        }

        shape := FindNewFreeform(slide.Shapes, nativeFreeform.originalIds)
        if !shape
            return
        FinishNativeFreeform(shape, nativeFreeform.hwnd)
    }
}

FinishNativeFreeform(shape, hwnd) {
    Critical
    StopTrackingFreeform()
    if WinActive("ahk_id " hwnd) {
        ; Release PowerPoint's drawing tool before using shape shortcuts.
        SendEvent "{Esc}"
        Sleep 50
    }
    try StyleOutline(shape)
    try shape.Line.BeginArrowheadStyle := 1 ; msoArrowheadNone
    try shape.Line.EndArrowheadStyle := 1
    if WinActive("ahk_id " hwnd)
        try shape.Select()
}

FindNewFreeform(shapes, originalIds) {
    count := shapes.Count
    if count <= originalIds.Count
        return false
    Loop count {
        shape := shapes.Item(count - A_Index + 1)
        if (shape.Type = 5 || shape.Type = 9) && !originalIds.Has(shape.Id)
            return shape
    }
    return false
}

StopTrackingFreeform() {
    global nativeFreeform
    SetTimer TrackNativeFreeform, 0
    nativeFreeform := false
}

CycleSelectedArrowheads(style := 2) {
    TrackNativeFreeform()
    changed := false
    try {
        ppt := ComObjActive("PowerPoint.Application")
        selection := ppt.ActiveWindow.Selection
        if selection.Type != 2 ; ppSelectionShapes
            return false
        selected := selection.ShapeRange
        Loop selected.Count {
            shape := selected.Item(A_Index)
            if shape.Type = 5 || shape.Type = 9 { ; msoFreeform or msoLine (including connectors)
                CycleArrowheads(shape.Line, style)
                changed := true
            }
        }
    }
    return changed
}

CycleArrowheads(line, style := 2) {
    begin := line.BeginArrowheadStyle != 1
    end := line.EndArrowheadStyle != 1
    next := Mod(begin + 2 * end + 1, 4)
    line.BeginArrowheadStyle := (next & 1) ? style : 1
    line.EndArrowheadStyle := (next & 2) ? style : 1
}

CycleSelectedDashStyles() {
    ApplyToSelectedShapes((shape) => shape.Line.DashStyle := NextDashStyle(shape.Line.DashStyle))
}

AdjustSelectedOutlineWeight(delta) {
    ApplyToSelectedShapes((shape) => shape.Line.Weight := AdjustedOutlineWeight(shape.Line.Weight, delta))
}

AdjustedOutlineWeight(current, delta) {
    return Max(0.25, current + delta)
}

ToggleShapeGlow(shape) {
    glow := shape.Glow
    if glow.Radius > 0 {
        glow.Radius := 0
        return
    }
    ApplyWhiteGlow(glow)
}

ApplyWhiteGlow(glow, radius := 3, transparency := 0.10) {
    glow.Color.RGB := 0xFFFFFF
    glow.Transparency := transparency
    glow.Radius := radius
}

ApplyToSelectedShapes(changeShape) {
    TrackNativeFreeform()
    try {
        ppt := ComObjActive("PowerPoint.Application")
        selection := ppt.ActiveWindow.Selection
        if selection.Type != 2 ; ppSelectionShapes
            return
        selected := selection.ShapeRange
        Loop selected.Count {
            try changeShape(selected.Item(A_Index))
        }
    }
}

NextDashStyle(current) {
    switch current {
        case 1: return 2 ; solid to square dot
        case 2: return 4 ; square dot to dash
        default: return 1
    }
}

CycleSelectedColors() {
    TrackNativeFreeform()
    changed := false
    try {
        ppt := ComObjActive("PowerPoint.Application")
        selection := ppt.ActiveWindow.Selection
        if selection.Type != 2 ; ppSelectionShapes
            return false
        selected := selection.ShapeRange
        Loop selected.Count {
            try {
                if CycleShapeColor(selected.Item(A_Index))
                    changed := true
            }
        }
    }
    return changed
}

CycleShapeColor(shape) {
    if shape.Type = 17 { ; msoTextBox
        color := shape.TextFrame.TextRange.Font.Color
    } else if shape.Tags.Item("PPT_RED_BOX_DRAWING") = "1" {
        color := shape.Line.ForeColor
    } else {
        return false
    }
    color.RGB := NextOutlineColor(color.RGB)
    return true
}

NextOutlineColor(current) {
    for index, color in outlineColors {
        if current = color
            return outlineColors[Mod(index, outlineColors.Length) + 1]
    }
    return outlineColors[1]
}

StyleOutline(shape) {
    shape.Fill.Visible := 0
    shape.Line.Visible := -1
    shape.Line.ForeColor.RGB := outlineColors[1]
    shape.Line.Weight := outlineWeight
    shape.Line.DashStyle := 1 ; msoLineSolid
    shape.Tags.Add("PPT_RED_BOX_DRAWING", "1")
    try ApplyWhiteGlow(shape.Glow)
}

ShowHint(message) {
    ToolTip message
    SetTimer HideHint, -2200
}

HideHint() {
    ToolTip()
}

InitializeApp() {
    LoadFontSettings()
    A_IconTip := "PowerPoint 그리기 도구"
    A_TrayMenu.Add("단축키 목록", ShowShortcutWindow)
    A_TrayMenu.Default := "단축키 목록"
    A_TrayMenu.ClickCount := 2
    A_TrayMenu.Add("업데이트 확인", (*) => CheckForUpdate(true))
    ShowShortcutWindow()
    SetTimer CheckForUpdate, -3000
    SetTimer CheckForUpdatePeriodically, 3600000
}

ShowShortcutWindow(*) {
    global updateStatus
    static window := 0
    if !window {
        window := Gui(, "PowerPoint 그리기 도구")
        window.MarginX := 18, window.MarginY := 16
        window.SetFont("s13 Bold", "Malgun Gothic")
        tabs := window.AddTab3("w550 h310", ["PowerPoint", "CATIA V5"])
        tabs.UseTab(1)
        window.AddText("w515", "PowerPoint 단축키")
        window.SetFont("s9 Norm", "Malgun Gothic")
        list := window.AddListView("w515 r9 -Multi", ["단축키", "동작"])
        for shortcut in [
            ["휠 클릭", "PowerPoint 자유형 시작"],
            ["휠 드래그", "빨간 사각형 그리기"],
            ["Ctrl+Shift+A", "화살촉 → 시작점, 끝점, 양쪽, 없음"],
            ["Ctrl+Shift+Q", "도트 → 시작점, 끝점, 양쪽, 없음"],
            ["Ctrl+Shift+C", "선·글자색 → 빨강, 파랑, 노랑, 초록, 검정"],
            ["Ctrl+Shift+N", "흰색 네온 효과 켜기·끄기 (3pt, 투명도 10%)"],
            ["F3", "선택한 텍스트에 기본 글꼴·크기 적용"],
            ["Ctrl+1", "5행 3열 표 · 기본 글꼴·크기, 가운데 정렬"],
            ["Shift+휠", "표 행 → 아래로 추가, 위로 마지막 행 삭제"],
            ["Ctrl+Shift+X", "텍스트 상자 → 가운데 정렬, 자동 크기, 글자 네온"],
            ["Ctrl+Shift+" Chr(96), "선 종류 → 실선, 사각 점선, 파선"],
            ["Ctrl+Shift+휠", "표 열 추가·삭제 / 도형 선 두께 조절"]
        ]
            list.Add(, shortcut[1], shortcut[2])
        list.ModifyCol(1, 125)
        list.ModifyCol(2, 382)
        tabs.UseTab(2)
        window.AddText("w515", "CATIA V5 · Part Design / GSD")
        catiaList := window.AddListView("w515 r9 -Multi", ["단축키", "동작"])
        catiaList.Add(, "MM", "Measure Between 실행")
        catiaList.Add(, "Ctrl+1", "모든 솔리드 Body 숨기기·표시")
        catiaList.Add(, "Ctrl+2", "모든 스케치 숨기기·표시")
        catiaList.Add(, "Ctrl+3", "3D 점·선·곡선 숨기기·표시 (교차·투영 포함)")
        catiaList.Add(, "Ctrl+Q", "새 Body 만들기")
        catiaList.Add(, "Ctrl+W", "새 Geometrical Set 만들기")
        catiaList.Add(, "Ctrl+R", "최상위 Part의 Reorder Children 열기")
        catiaList.Add(, "F", "스케치의 선택 개체 Fix 적용·해제")
        catiaList.Add(, "HH", "스케치: 보조선 전환 · 3D/트리: 숨김·표시+서브트리")
        catiaList.Add(, "PP", "스케치: Project 3D Elements · 스케치 밖: Plane")
        catiaList.Add(, "II", "스케치: Intersect 3D Elements · 스케치 밖: Isolate")
        catiaList.Add(, "Ctrl+T", "Thickness 실행")
        catiaList.Add(, "TT", "Thick Surface · 서피스에 두께를 줘 솔리드 생성")
        catiaList.Add(, "SS", "Part Design: 솔리드 Split · GSD: 서피스 Split")
        catiaList.Add(, "DD", "스케치에서 Point by Clicking 실행")
        catiaList.Add(, "LL", "스케치: 2D Line · 스케치 밖: GSD Line")
        catiaList.Add(, "VV", "스케치: Normal View · 3D: 현재 뷰 반대 방향")
        catiaList.Add(, "Ctrl+넘패드 5", "상단 등각도 네 방향 순환")
        catiaList.Add(, "Ctrl+넘패드 8/2", "시선 방향 −Z / +Z")
        catiaList.Add(, "Ctrl+넘패드 4/6", "시선 방향 +Y / −Y")
        catiaList.Add(, "Ctrl+넘패드 7/9/1/3", "인접한 두 방향 사이의 45도 뷰")
        catiaList.Add(, "CC", "스케치의 선택한 두 개체 Coincidence 적용")
        catiaList.Add(, "QQ", "스케치 Quick Trim 연속 실행 · Esc로 종료")
        catiaList.Add(, "EE", "GSD에서 Multiple Extract 실행")
        catiaList.Add(, "XX", "GSD에서 Extrapolate 실행")
        catiaList.Add(, "JJ", "GSD에서 Join 실행")
        catiaList.Add(, "Ctrl+E", "Extruded Surface Definition 열기")
        catiaList.Add(, "RR", "Remove Face 연속 실행 · Esc로 종료")
        catiaList.Add(, "Ctrl+F", "상위 Body·Geometrical Set을 Define In Work Object")
        catiaList.Add(, "FF", "GSD에서 Fill 실행")
        catiaList.ModifyCol(1, 125)
        catiaList.ModifyCol(2, 382)
        window.AddText("w515", "현재 CATPart 기준 · 표시된 Body가 있으면 모두 숨기고, 모두 숨겨져 있으면 표시")
        tabs.UseTab()
        window.AddText("xm y+14", "기본 글꼴")
        fontChoice := window.AddDropDownList("x+8 yp-3 w190 Choose1", ["현대하모니 M", "나눔바른고딕 OTF"])
        fontChoice.Choose(defaultFontName = "현대하모니 M" ? 1 : 2)
        window.AddText("x+12 yp+3", "크기(pt)")
        sizeChoice := window.AddComboBox("x+8 yp-3 w70", ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "36", "48", "60", "72"])
        sizeChoice.Text := defaultFontSize
        saveFont := window.AddButton("x+8 yp w55", "저장")
        fontStatus := window.AddText("xm y+7 w515", "텍스트 상자·표 생성과 F3에 적용됩니다.")
        saveFont.OnEvent("Click", SaveFontSettings.Bind(fontChoice, sizeChoice, fontStatus))
        window.AddText("xm y+14", "자주 쓰는 기호")
        copyStatus := window.AddText("x+10 yp w250", "")
        window.SetFont("s13", "Segoe UI Symbol")
        for index, symbol in ["⌀", "↑", "↓", "←", "→"] {
            position := index = 1 ? "xm y+6" : "x+6 yp"
            button := window.AddButton(position " w48 h32", symbol)
            button.OnEvent("Click", CopySymbol.Bind(symbol, copyStatus))
        }
        window.SetFont("s9", "Malgun Gothic")
        window.AddText("xm y+14 w515", "버전 " appVersion)
        updateStatus := window.AddText("xm y+5 w515", "실행 중에 자동으로 업데이트를 확인합니다.")
        window.AddButton("xm y+10 w120", "업데이트 확인").OnEvent("Click", (*) => CheckForUpdate(true))
        window.OnEvent("Close", (gui) => gui.Hide())
        window.OnEvent("Escape", (gui) => gui.Hide())
    }
    window.Show()
}

CopySymbol(symbol, status, *) {
    A_Clipboard := symbol
    status.Text := symbol " 복사됨"
}

CheckForUpdatePeriodically() {
    CheckForUpdate()
}

CheckForUpdate(manual := false) {
    global updateRequest, updateDeadline
    if updateRequest || pendingUpdate != ""
        return
    try {
        request := ComObject("WinHttp.WinHttpRequest.5.1")
        request.SetTimeouts(5000, 5000, 5000, 8000)
        request.Open("GET", updateUrl, true)
        request.SetRequestHeader("Cache-Control", "no-cache")
        request.Send()
        updateRequest := request
        updateDeadline := A_TickCount + 25000
        SetTimer ReadUpdate, 250
        if manual
            SetUpdateStatus("업데이트를 확인하는 중입니다.")
    } catch {
        SetUpdateStatus("업데이트 서버에 연결할 수 없습니다.")
    }
}

ReadUpdate() {
    global updateRequest, pendingUpdate
    try {
        if !updateRequest.WaitForResponse(0) {
            if A_TickCount >= updateDeadline {
                updateRequest.Abort()
                throw Error("Update request timed out.")
            }
            return
        }
        if updateRequest.Status != 200
            throw Error("Update response was not successful.")
        source := updateRequest.ResponseText
        version := ReadScriptVersion(source)
        if VerCompare(version, appVersion) > 0 {
            pendingUpdate := source
            SetUpdateStatus("새 버전 " version "을 받았습니다. PowerPoint 창이 비활성화되면 적용합니다.")
            SetTimer ApplyUpdateWhenIdle, 1000
        } else {
            SetUpdateStatus("최신 버전을 사용 중입니다.")
        }
    } catch {
        SetUpdateStatus("업데이트를 확인할 수 없습니다.")
    }
    SetTimer ReadUpdate, 0
    updateRequest := false
}

ReadScriptVersion(source) {
    if !RegExMatch(source, 'm)^appVersion := "(\d+\.\d+\.\d+)"\r?$', &version)
        throw Error("Update version is missing.")
    return version[1]
}

ApplyUpdateWhenIdle() {
    global pendingUpdate
    if WinActive("ahk_exe POWERPNT.EXE") || GetKeyState("MButton", "P")
        return

    SetTimer ApplyUpdateWhenIdle, 0
    staged := A_ScriptFullPath "." A_TickCount ".update.ahk"
    try {
        file := FileOpen(staged, "w", "UTF-8")
        try file.Write(LTrim(pendingUpdate, Chr(0xFEFF)))
        finally file.Close()
        validation := Format('"{}" /ErrorStdOut /validate "{}"', A_AhkPath, staged)
        if RunWait(validation, , "Hide") != 0
            throw Error("Downloaded script failed validation.")
        if WinActive("ahk_exe POWERPNT.EXE") || GetKeyState("MButton", "P") {
            SetTimer ApplyUpdateWhenIdle, 1000
            return
        }
        if !DllCall("MoveFileExW", "Str", staged, "Str", A_ScriptFullPath, "UInt", 9)
            throw OSError()
        pendingUpdate := ""
        SetUpdateStatus("업데이트를 적용했습니다. 앱을 다시 시작합니다.")
        Reload()
        SetTimer ReportReloadFailure, -5000
    } catch {
        pendingUpdate := ""
        SetUpdateStatus("다운로드한 업데이트를 적용할 수 없습니다.")
    } finally {
        if FileExist(staged)
            try FileDelete(staged)
    }
}

ReportReloadFailure() {
    SetUpdateStatus("새 버전을 저장했습니다. 적용하려면 이 파일을 다시 실행하세요.")
}

SetUpdateStatus(message) {
    if updateStatus
        updateStatus.Text := message
}
