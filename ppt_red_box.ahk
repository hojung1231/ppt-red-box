#Requires AutoHotkey v2.0
#SingleInstance Force
CoordMode "Mouse", "Screen"

appVersion := "1.0.7"
updateUrl := "https://raw.githubusercontent.com/hojung1231/ppt-red-box/main/ppt_red_box.ahk"
updateRequest := false
updateDeadline := 0
pendingUpdate := ""
updateStatus := ""
outlineWeight := 1.5
outlineColors := [0x0000FF, 0xFF0000, 0x00FFFF, 0x00FF00, 0x000000] ; Red, blue, yellow, green, black in Office BGR
nativeFreeform := false

if A_LineFile = A_ScriptFullPath
    InitializeApp()

#HotIf WinActive("ahk_exe POWERPNT.EXE")
F3::ApplyNanumFontToSelection()
^+x::InsertTextbox()
^+n:: {
    ApplyToSelectedShapes(ToggleShapeGlow)
    KeyWait "n"
}
^+SC029::CycleSelectedDashStyles()
^+WheelUp::AdjustSelectedOutlineWeight(0.25)
^+WheelDown::AdjustSelectedOutlineWeight(-0.25)
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
    pointsPerCentimeter := 72 / 2.54
    width := 4 * pointsPerCentimeter
    height := 0.64 * pointsPerCentimeter
    left := Max(0, Min(point.x, slideWidth - width))
    top := Max(0, Min(point.y, slideHeight - height))
    shape := shapes.AddTextbox(1, left, top, width, height)
    try {
        fill := shape.Fill
        fill.Solid()
        fill.Visible := -1
        fill.ForeColor.RGB := 0xFFFFFF
        fill.Transparency := 0.30

        frame := shape.TextFrame
        frame.AutoSize := 0
        frame.WordWrap := -1
        shape.Width := width
        shape.Height := height
        font := frame.TextRange.Font
        SetNanumFont(font)
        font.Size := 12
        font.Color.RGB := 0x000000
    } catch as error {
        shape.Delete()
        throw error
    }
    return shape
}

ApplyNanumFontToSelection() {
    TrackNativeFreeform()
    try {
        ppt := ComObjActive("PowerPoint.Application")
        selection := ppt.ActiveWindow.Selection
        switch selection.Type {
            case 3: ; ppSelectionText
                SetNanumFont(selection.TextRange.Font)
            case 2: ; ppSelectionShapes
                selected := selection.ShapeRange
                Loop selected.Count {
                    shape := selected.Item(A_Index)
                    try {
                        if shape.HasTextFrame = -1
                            SetNanumFont(shape.TextFrame.TextRange.Font)
                    }
                }
        }
    }
}

SetNanumFont(font) {
    font.Name := "NanumBarunGothicOTF"
    font.NameFarEast := "NanumBarunGothicOTF"
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

CycleSelectedArrowheads() {
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
                CycleArrowheads(shape.Line)
                changed := true
            }
        }
    }
    return changed
}

CycleArrowheads(line) {
    begin := line.BeginArrowheadStyle != 1
    end := line.EndArrowheadStyle != 1
    next := Mod(begin + 2 * end + 1, 4)
    line.BeginArrowheadStyle := (next & 1) ? 2 : 1 ; triangle or none
    line.EndArrowheadStyle := (next & 2) ? 2 : 1
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

ApplyWhiteGlow(glow) {
    glow.Color.RGB := 0xFFFFFF
    glow.Transparency := 0.10
    glow.Radius := 3
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
        window.AddText("w515", "PowerPoint 단축키")
        window.SetFont("s9 Norm", "Malgun Gothic")
        list := window.AddListView("w515 r9 -Multi", ["단축키", "동작"])
        for shortcut in [
            ["휠 클릭", "PowerPoint 자유형 시작"],
            ["휠 드래그", "빨간 사각형 그리기"],
            ["Ctrl+Shift+A", "화살촉 → 시작점, 끝점, 양쪽, 없음"],
            ["Ctrl+Shift+C", "선·글자색 → 빨강, 파랑, 노랑, 초록, 검정"],
            ["Ctrl+Shift+N", "흰색 네온 효과 켜기·끄기 (3pt, 투명도 10%)"],
            ["F3", "선택한 텍스트 글꼴 → 나눔바른고딕 OTF"],
            ["Ctrl+Shift+X", "텍스트 상자 → 4 × 0.64cm, 나눔바른고딕 OTF 12pt"],
            ["Ctrl+Shift+" Chr(96), "선 종류 → 실선, 사각 점선, 파선"],
            ["Ctrl+Shift+휠", "선 두께 → 위로 0.25pt 증가, 아래로 0.25pt 감소"]
        ]
            list.Add(, shortcut[1], shortcut[2])
        list.ModifyCol(1, 125)
        list.ModifyCol(2, 382)
        window.AddText("xm y+14 w515", "버전 " appVersion)
        updateStatus := window.AddText("xm y+5 w515", "실행 중에 자동으로 업데이트를 확인합니다.")
        window.AddButton("xm y+10 w120", "업데이트 확인").OnEvent("Click", (*) => CheckForUpdate(true))
        window.OnEvent("Close", (gui) => gui.Hide())
        window.OnEvent("Escape", (gui) => gui.Hide())
    }
    window.Show()
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
