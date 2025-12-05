#Requires AutoHotkey v2.0-a
#SingleInstance


;============================================================================================
; @Dependencies
;============================================================================================
#Include ..\..\lib\Gdip_All.ahk


;============================================================================================
; @Auto_Execute
;============================================================================================
pToken := Gdip_Startup()
OnExit((*) => Gdip_Shutdown(pToken))


;============================================================================================
; @TaskSwitcher
;============================================================================================
class TaskSwitcher {
    ; @options that can be changed here or used as a property name when passing options to TaskSwitcher({option: value})
    ; Note - options that go through Gdip_TextToGraphics require ARGB format as a string (e.g. 'FF00FF00' is green)
    ;       while other color options use 0xARGB as a hex number (e.g. 0xFF00FF00 is green)
    static defaultTextColor := 0xFFFFFFFF
    static highlightTextColor := 0xFF6995DB
    static rowHighlightColor := 0x30FFFFFF
    static mouseHighlightTextColor := this.highlightTextColor
    static backgroundColor := 0xFF333333
    static dividerColor := 0xFFFFFFFF
    static bannerColor := 0xFF1B56B5
    static bannerTextColor := 'FFFFFFFF'
    static bannerText := 'Task Switcher'
    static wrapRowSelection := true
    static alwaysHighlightFirst := false
    static defaultSearchText := ''
    static searchTextColor := 'ffc8c8c8'
    static searchBackgroundColor := this.searchTextColor
    static escapeAlwaysClose := false


    static isOpen => WinExist('ahk_id' this.menu.Hwnd)
    static isActive => WinActive('ahk_id' this.menu.Hwnd)
    static hasMouseOver => (MouseGetPos(,, &win), win = TaskSwitcher.menu.Hwnd)

    static Call(options := {}) {
        for option, value in options.OwnProps() {
            this.%option% := value
        }
    }

    ; sorts the windows alphabetically
    static OpenMenuSorted() {
        this.OpenMenu(true)
    }

    static ToggleMenuSorted() {
        this.ToggleMenu(true)
    }

    static ToggleMenu(sortedWindows := false) {
        if WinExist('ahk_id ' this.menu.Hwnd) {
            this.CloseMenu()
            return
        }

        this.OpenMenu(sortedWindows)
    }

    ; uses z-order of windows
    static OpenMenu(sortedWindows := false) {
        if WinExist('ahk_id ' this.menu.Hwnd) {
            return
        }

        this._sortedWindows := sortedWindows

        ; setup message handlers
        OnMessage(0x200, this._OnMouseMove)
        OnMessage(0x20A, this._OnMouseWheel)
        OnMessage(0x2A3, this._OnMouseLeave)
        OnMessage(0x201, this._OnLeftClick)
        OnMessage(0x202, this._OnLeftClickRelease)

        this.__RefreshWindows()
        this.__CreateMenu()
        this._ih.Start()
    }

    static CloseMenu() {
        if !WinExist('ahk_id ' this.menu.Hwnd) {
            return
        }

        this._ih.Stop()
        this.menu.Hide()
        OnMessage(0x200, this._OnMouseMove, 0)
        OnMessage(0x20A, this._OnMouseWheel, 0)
        OnMessage(0x2A3, this._OnMouseLeave, 0)
        OnMessage(0x201, this._OnLeftClick, 0)
        OnMessage(0x202, this._OnLeftClickRelease, 0)

        if this._scrollTimer {
            SetTimer(this._scrollTimer, 0)
            this._scrollTimer := 0
        }

        this.__GDIP_Cleanup()
    }

    static ActivateWindowAndCloseMenu() {
        this.CloseMenu()
        this.ActivateWindow()
    }

    static ActivateWindow() {
        window := this.menu.windows[this._selectedRow]
            this._onWindowActivate(window)
        if this.__ActivateWindow(window) {
        }
    }

    static SelectPreviousWindow() {
        if this._selectedRow > 1 {
            this._selectedRow -= 1
            this.__ScrollToSelectedRow()
        } else if this.wrapRowSelection {
            this._selectedRow := this.menu.windows.Length
            this.__ScrollToSelectedRow()
        }

        this.__DrawMenu()
    }

    static SelectNextWindow() {
        if this._selectedRow < this.menu.windows.Length {
            this._selectedRow += 1
            this.__ScrollToSelectedRow()
        } else if this.wrapRowSelection {
            this._selectedRow := 1
            this.__ScrollToSelectedRow()
        }

        this.__DrawMenu()
    }

    static OnWindowActivate(Callback) {
        this._onWindowActivate := (_, params*) => Callback(params*)
    }

    ; pass 'Off' if you ever want to disable those hotkeys after already being active
    ; pass 'Toggle' if you want to toggle the state
    static AltTabReplacement(state := 'On') {
        static altTabHotkeysEnabled := false,
            previousState := state

        if state = 'Toggle' {
            state := previousState = 'On' ? 'Off' : 'On'
        }

        HotIf((*) => !TaskSwitcher.isActive)
        Hotkey('!Tab', (*) {
            altTabHotkeysEnabled := true
            TaskSwitcher.OpenMenu()
            TaskSwitcher.SelectNextWindow()
        }, state)

        HotIf((*) => TaskSwitcher.isActive)
        Hotkey('!Tab', (*) => TaskSwitcher.SelectNextWindow(), state)
        Hotkey('+!Tab', (*) => TaskSwitcher.SelectPreviousWindow(), state)
        HotIf((*) => altTabHotkeysEnabled)
        Hotkey('~*Alt up', (*) {
            ; prevents alt release from closing window if it wasn't opened with alt-tab method
            ; this is in case anyone uses hotkeys that allow something like alt+up/down to navigate, the menu will not close when releasing alt
            if altTabHotkeysEnabled {
                TaskSwitcher.ActivateWindowAndCloseMenu()
                altTabHotkeysEnabled := false
            }
        }, state)

        previousState := state
    }

    static __CreateMenu() {
        this._windowRects := []
        this._selectedRow := 1
        this._mousedOver := 0
        this._leftClicked := 0
        this._searchText := this.defaultSearchText
        this._mouseLeft := true

        this._scrollOffset := 0
        this._targetScrollOffset := 0    ; reset target
        this._scrollTimer := 0           ; reset timer

        totalHeight := this.__CalculateTotalHeight()
        this.__Init_GDIP(totalHeight)

        ; initial draw
        this.__DrawMenu()

        ; setup window
        this.menu.Show('Hide w' this._maxWidth ' h' totalHeight)
        this.menu.GetPos(&winX, &winY)

        ; center the window manually
        MonitorGetWorkArea(, &left, &top, &right, &bottom)
        centerX := left + (right - left - this._maxWidth) / 2
        centerY := top + (bottom - top - totalHeight) / 2

        FrameShadow(this.menu.Hwnd)
        UpdateLayeredWindow(this.menu.Hwnd, this._hdc, centerX, centerY, this._maxWidth, totalHeight)
        this.menu.Show()

        FrameShadow(hwnd) {
            DllCall('dwmapi.dll\DwmIsCompositionEnabled', 'Int*', &dwmEnabled:=0)

            if !dwmEnabled {
                DllCall('user32.dll\SetClassLongPtr', 'Ptr', hwnd, 'Int', -26, 'Ptr', DllCall('user32.dll\GetClassLongPtr', 'Ptr', hwnd, 'Int', -26) | 0x20000)
                return
            }

            NumPut('Int', 1, 'Int', 1, 'Int', 1, 'Int', 1, margins := Buffer(16, 0))
            DllCall('dwmapi.dll\DwmSetWindowAttribute', 'Ptr', hwnd, 'Int', 33, 'Int*', 2, 'Int', 4)
            DllCall('dwmapi.dll\DwmExtendFrameIntoClientArea', 'Ptr', hwnd, 'Ptr', margins)
        }
    }

    static __CloseWindow(window) {
        this._ih.Stop()
        WinClose(window.hwnd)
        WinWaitClose(window.hwnd)
        this._ih.Start()

        this.__RefreshWindows()
        this.__ApplySearchFilter()
        this.__UpdateMenuWhenFiltered()
    }

    static __ApplySearchFilter() {
        if this._searchText != this.defaultSearchText && StrLen(this._searchText) > 0 {
            matches := []
            for win in this._allWindows {
                if InStr(win.name, this._searchText) || InStr(win.title, this._searchText) {
                    matches.Push(win)
                }
            }
            this.menu.windows := matches
        } else {
            this.menu.windows := this._allWindows.Clone()
        }
    }

    static __CalculateTotalHeight() {
        totalRows := this.menu.windows.Length

        if totalRows > this._maxVisibleRows {
            ; show partial row to indicate scrollability
            visibleRows := this._maxVisibleRows - 0.5  ; Show half of the next row
            totalDividers := Floor(visibleRows)
            contentHeight := Integer(visibleRows * this._rowHeight) + (totalDividers * this._dividerHeight)
        } else {
            totalDividers := Max(0, totalRows - 1)
            contentHeight := (totalRows * this._rowHeight) + (totalDividers * this._dividerHeight)
        }

        totalHeight := this._bannerHeight + contentHeight
        return totalHeight
    }

    static __DrawMenu() {
        totalHeight := this.__CalculateTotalHeight()

        ; clear background
        pBrush := Gdip_BrushCreateSolid(this.backgroundColor)
        Gdip_FillRectangle(this._pGraphics, pBrush, 0, 0, this._maxWidth, totalHeight)
        Gdip_DeleteBrush(pBrush)

        ; draw banner
        pBrushBanner := Gdip_BrushCreateSolid(this.bannerColor)
        Gdip_FillRectangle(this._pGraphics, pBrushBanner, 0, 0, this._maxWidth, this._bannerHeight)
        Gdip_DeleteBrush(pBrushBanner)

        ; draw banner text
        options := 'x' this._marginX ' y16 s18 Bold c' this.bannerTextColor
        Gdip_TextToGraphics(this._pGraphics, this.bannerText, options, 'Arial', this._maxWidth - (this._marginX * 2), this._bannerHeight)

        if this.searchBackgroundColor != this.searchTextColor {
            rect := {
                x: this._maxWidth//2 + this._marginX,
                y: 8,
                w: this._maxWidth//2 - (this._marginX * 2) + 4,
                h: 34,
                r: 8
            }
            this.searchBackgroundRect := rect
            pBrushDebug := Gdip_BrushCreateSolid(this.searchBackgroundColor)
            Gdip_FillRoundedRectangle(this._pGraphics, pBrushDebug, rect.x, rect.y, rect.w, rect.h, rect.r)
            Gdip_DeleteBrush(pBrushDebug)
        }

        ; draw input text (right-aligned)
        displayText := this._searchText . Chr(0x200B)
        inputOptions := 'x' (this._maxWidth - 380) ' y16 Right'
        inputOptions .= (this._searchText = this.defaultSearchText)
            ? 's16 Italic c' this.searchTextColor
            : 's18 Bold c' this.bannerTextColor
        Gdip_TextToGraphics(this._pGraphics, displayText, inputOptions, 'Arial', 370, this._bannerHeight)

        ; set clipping - only draw content below banner
        Gdip_SetClipRect(this._pGraphics, 0, this._bannerHeight, this._maxWidth, totalHeight - this._bannerHeight)

        ; draw window rows with pixel offset
        rowWithDivider := this._rowHeight + this._dividerHeight
        this._windowRects := []
        this._closeButtonRects := []  ; Track close button positions

        for index, window in this.menu.windows {
            ; calculate y position with scroll offset
            rowY := this._bannerHeight + ((index - 1) * rowWithDivider) - this._scrollOffset

            ; skip only if COMPLETELY outside visible area
            if rowY + this._rowHeight < this._bannerHeight || rowY > totalHeight {
                continue
            }

            this._windowRects.Push({
                x: 0,
                y: Max(rowY, this._bannerHeight),
                w: this._maxWidth,
                h: this._rowHeight,
                window: window,
                actualIndex: index
            })

            ; highlight selected row
            if this._selectedRow = index {
                pBrushHover := Gdip_BrushCreateSolid(this.rowHighlightColor)
                Gdip_FillRectangle(this._pGraphics, pBrushHover, 0, rowY, this._maxWidth, this._rowHeight)
                Gdip_DeleteBrush(pBrushHover)
            }

            ; draw icon
            iconX := this._marginX
            iconY := rowY + (this._rowHeight - this._iconSize) / 2
            this.__DrawIcon(window, iconX, iconY)

            ; draw row
            switch index {
            case this._selectedRow:
                textColor := this.highlightTextColor
            case this._mousedOver:
                textColor := this.mouseHighlightTextColor
            default:
                textColor := this.defaultTextColor
            }

            titleX := iconX + this._iconSize + 15
            windowOptions := 'x' titleX ' y' (rowY + 8) ' s18 Bold cFF' SubStr(Format('{:06X}', textColor), 3)
            titleOptions := 'x' titleX ' y' (rowY + 28) ' s16 cFF' SubStr(Format('{:06X}', textColor), 3)
            Gdip_TextToGraphics(this._pGraphics, window.name, windowOptions, 'Arial', this._maxWidth - titleX - this._marginX - 40, this._rowHeight)
            Gdip_TextToGraphics(this._pGraphics, window.title, titleOptions, 'Arial', this._maxWidth - titleX - this._marginX - 40, this._rowHeight)

            ; draw close button (X)
            closeButtonSize := 24
            closeButtonX := this._maxWidth - this._marginX - closeButtonSize - 10
            closeButtonY := rowY + (this._rowHeight - closeButtonSize) / 2

            this._closeButtonRects.Push({
                x: closeButtonX,
                y: closeButtonY,
                w: closeButtonSize,
                h: closeButtonSize,
                actualIndex: index
            })

            ; check if mouse is over this close button
            isHoveringCloseButton := (this._hoveredCloseButton = index)

            ; draw close button background (highlight if hovering)
            if isHoveringCloseButton {
                pBrushClose := Gdip_BrushCreateSolid(0x80FF0000)  ; Semi-transparent red
            } else {
                pBrushClose := Gdip_BrushCreateSolid(0x40FFFFFF)  ; Semi-transparent white
            }
            Gdip_FillEllipse(this._pGraphics, pBrushClose, closeButtonX, closeButtonY, closeButtonSize, closeButtonSize)
            Gdip_DeleteBrush(pBrushClose)

            ; draw X
            pPen := Gdip_CreatePen(isHoveringCloseButton ? 0xFFFFFFFF : 0xFFAAAAAA, 2)
            offset := 6
            Gdip_DrawLine(this._pGraphics, pPen, closeButtonX + offset, closeButtonY + offset, closeButtonX + closeButtonSize - offset, closeButtonY + closeButtonSize - offset)
            Gdip_DrawLine(this._pGraphics, pPen, closeButtonX + closeButtonSize - offset, closeButtonY + offset, closeButtonX + offset, closeButtonY + closeButtonSize - offset)
            Gdip_DeletePen(pPen)

            ; draw divider (skip if outside visible area or last item)
            if index < this.menu.windows.Length {
                dividerY := rowY + this._rowHeight
                if dividerY > this._bannerHeight && dividerY < totalHeight {
                    pBrushDiv := Gdip_BrushCreateSolid(this.dividerColor)
                    Gdip_FillRectangle(this._pGraphics, pBrushDiv, this._marginX, dividerY, this._maxWidth - (this._marginX * 2), this._dividerHeight)
                    Gdip_DeleteBrush(pBrushDiv)
                }
            }
        }

        ; reset clipping
        Gdip_ResetClip(this._pGraphics)

        ; get current window position to maintain it
        this.menu.GetPos(&winX, &winY)
        UpdateLayeredWindow(this.menu.Hwnd, this._hdc, winX, winY, this._maxWidth, totalHeight)
    }

    static __UpdateMenuWhenFiltered() {
        ; decide which row to highlight
        totalRows := this.menu.windows.Length
        if this.alwaysHighlightFirst {
            this._selectedRow := 1
        } else if this._selectedRow > totalRows {
            this._selectedRow := Max(1, totalRows)
        }

        this._scrollOffset := 0

        this.__GDIP_Cleanup()
        totalHeight := this.__CalculateTotalHeight()
        this.__Init_GDIP(totalHeight)
        this.__DrawMenu()

        ; recenter vertically if height changed
        this.menu.GetPos(&winX, &winY)
        MonitorGetWorkArea(, &left, &top, &right, &bottom)
        centerY := top + (bottom - top - totalHeight) / 2

        UpdateLayeredWindow(this.menu.Hwnd, this._hdc, winX, centerY, this._maxWidth, totalHeight)
    }

    static __RefreshWindows() {
        list := this.__AltTabWindows()
        windows := []

        for id in list {
            processName := this.__GetWindowName(WinGetProcessPath(id))
            if processName = '' {
                processName := StrSplit(WinGetProcessName(id), '.exe')[1]
            }
            processTitle := WinGetTitle(id)

            windows.Push({
                hwnd:   id,
                title:  processTitle,
                name:   processName,
            })
        }

        if this._sortedWindows {
            SortWindows()
        }

        this._allWindows := windows.Clone()  ; store complete list
        this.menu.windows := windows

        SortWindows() {
            i := 2
            while i <= windows.Length {
                temp := windows[i]
                j := i - 1
                while j >= 1 and StrCompare(windows[j].name, temp.name) > 0 {
                    windows[j + 1] := windows[j]
                    j--
                }
                windows[j + 1] := temp
                i++
            }
        }
    }

    static __Init_GDIP(totalHeight) {
        ; create GDI+ bitmap and graphics
        this._hBitmap := CreateDIBSection(this._maxWidth, totalHeight)
        this._hdc := CreateCompatibleDC()
        this.obm := SelectObject(this._hdc, this._hBitmap)
        this._pGraphics := Gdip_GraphicsFromHDC(this._hdc)
        Gdip_SetSmoothingMode(this._pGraphics, 4)
        Gdip_SetTextRenderingHint(this._pGraphics, 3)
    }

    static __GetWindowName(path) {
        size := DllCall('version\GetFileVersionInfoSizeW', 'str', path, 'uint*', 0, 'uint')
        if !size {
            return
        }

        buf := Buffer(size)
        if !DllCall('version\GetFileVersionInfoW', 'str', path, 'uint', 0, 'uint', size, 'ptr', buf) {
            return
        }

        Query(val) {
            ptr := 0, len := 0
            if DllCall('version\VerQueryValueW',
                'ptr', buf,
                'str', '\StringFileInfo\040904b0\' val,
                'ptr*', &ptr,
                'uint*', &len)
            {
                return StrGet(ptr, 'UTF-16')
            }
        }

        return Query('ProductName')
    }

    static __DrawIcon(window, x, y) {
        hwnd := window.hwnd
        path := WinGetProcessPath(hwnd)

        ; use actual app path for cache key, not ApplicationFrameHost
        cacheKey := path

        ; for ApplicationFrameHost, try to get the real app
        if InStr(path, "ApplicationFrameHost.exe") {
            try {
                uwpPath := this.__GetLargestUWPLogoPath(hwnd)
                if uwpPath {
                    cacheKey := uwpPath  ; use UWP path as cache key
                }
            }
        }

        if !this._iconCache.Has(cacheKey) {
            pBitmap := 0
            isUWP := false

            try {
                ; try UWP FIRST for WindowsApps or ApplicationFrameHost
                if InStr(path, "WindowsApps") || InStr(path, "ApplicationFrameHost") {
                    try {
                        uwpPath := this.__GetLargestUWPLogoPath(hwnd)
                        if uwpPath && FileExist(uwpPath) {
                            pBitmap := Gdip_CreateBitmapFromFile(uwpPath)
                            if pBitmap {
                                isUWP := true  ; Only mark as UWP if we actually got the bitmap
                            }
                        }
                    }
                }

                ; if no UWP icon, try regular extraction
                if !pBitmap {
                    for size in [256, 128, 48] {
                        hIcon := 0
                        DllCall('PrivateExtractIcons', 'Str', path, 'Int', 0, 'Int', size, 'Int', size, 'Ptr*', &hIcon, 'Ptr*', 0, 'UInt', 1, 'UInt', 0)

                        if hIcon {
                            pBitmap := Gdip_CreateBitmapFromHICON(hIcon)
                            DllCall('DestroyIcon', 'Ptr', hIcon)
                            break
                        }
                    }
                }
            }

            ; fallback to shell32 icon
            if !pBitmap {
                try {
                    hIcon := 0
                    DllCall("PrivateExtractIcons", "Str", "shell32.dll", "Int", 0, "Int", 48, "Int", 48, "Ptr*", &hIcon, "Ptr*", 0, "UInt", 1, "UInt", 0)

                    if hIcon {
                        pBitmap := Gdip_CreateBitmapFromHICON(hIcon)
                        DllCall("DestroyIcon", "Ptr", hIcon)
                    }
                }
            }

            this._iconCache[cacheKey] := {bitmap: pBitmap, isUWP: isUWP}
        }

        iconData := this._iconCache[cacheKey]
        pBitmap := iconData.bitmap

        if pBitmap && Gdip_GetImageWidth(pBitmap) {
            Gdip_SetInterpolationMode(this._pGraphics, 7)

            ; draw UWP icons slightly larger to compensate for smaller source images
            if iconData.isUWP {
                drawSize := this._iconSize * 1.25  ; 25% larger
                offset := (this._iconSize - drawSize) / 2  ; Center it
                Gdip_DrawImage(this._pGraphics, pBitmap, x + offset, y + offset, drawSize, drawSize)
            } else {
                Gdip_DrawImage(this._pGraphics, pBitmap, x, y, this._iconSize, this._iconSize)
            }

            Gdip_SetInterpolationMode(this._pGraphics, 2)
        }
    }

    static __GetLargestUWPLogoPath(hwnd) {
        Address := CallbackCreate(EnumChildProc.Bind(WinGetPID(hwnd)), 'Fast', 2)
        DllCall('User32.dll\EnumChildWindows', 'Ptr', hwnd, 'Ptr', Address, 'UInt*', &ChildPID := 0, 'Int')
        CallbackFree(Address)

        ; if no child PID, use the main window's PID
        if !ChildPID {
            ChildPID := WinGetPID(hwnd)
        }

        if !AppHasPackage(ChildPID) {
            return
        }

        try {
            processPath := ProcessGetPath(ChildPID)
            defaultLogoPath := GetDefaultLogoPath(processPath)
            largestPath := GetLargestLogoPath(defaultLogoPath)
            return largestPath
        } catch as e {
            return
        }

        EnumChildProc(PID, hwnd, lParam) {
            ChildPID := WinGetPID(hwnd)
            if ChildPID != PID {
                NumPut 'UInt', ChildPID, lParam
                return false
            }
            return true
        }

        AppHasPackage(ChildPID) {
            static PROCESS_QUERY_LIMITED_INFORMATION := 0x1000, APPMODEL_ERROR_NO_PACKAGE := 15700
            ProcessHandle := DllCall('Kernel32.dll\OpenProcess', 'UInt', PROCESS_QUERY_LIMITED_INFORMATION, 'Int', false, 'UInt', ChildPID, 'Ptr')
            IsUWP := DllCall('Kernel32.dll\GetPackageId', 'Ptr', ProcessHandle, 'UInt*', &BufferLength := 0, 'Ptr', 0, 'Int') != APPMODEL_ERROR_NO_PACKAGE
            DllCall('Kernel32.dll\CloseHandle', 'Ptr', ProcessHandle, 'Int')
            return IsUWP
        }

        GetDefaultLogoPath(Path) {
            SplitPath Path, , &Dir
            if !RegExMatch(FileRead(Dir '\AppxManifest.xml', 'UTF-8'), '<Logo>(.*)</Logo>', &Match) {
                throw Error('Unable to read logo information from file.', -1, Dir '\AppxManifest.xml')
            }
            return Dir '\' Match[1]
        }

        GetLargestLogoPath(Path) {
            LoopFileSize := 0
            SplitPath Path, , &Dir, &Extension, &NameNoExt
            Loop Files Dir '\' NameNoExt '.scale-*.' Extension {
                if A_LoopFileSize > LoopFileSize && RegExMatch(A_LoopFileName, '\d+\.' Extension '$') {
                    LoopFilePath := A_LoopFilePath, LoopFileSize := A_LoopFileSize
                }
            }
            return LoopFilePath ?? ''
        }
    }

    static __OnMouseMove(wParam, lParam, msg, hwnd) {
        static tme := TrackMouseLeave(hwnd)   ; required for OnMouseLeave to work correctly
        if this._mouseLeft {
            this._mouseLeft := false
            DllCall('user32.dll\TrackMouseEvent', 'Ptr', tme)
        }

        x := lParam & 0xFFFF
        y := lParam >> 16

        ; check if hovering over a close button first
        newHoveredCloseButton := 0
        for rect in this._closeButtonRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                newHoveredCloseButton := rect.actualIndex
                break
            }
        }

        ; check row hovers
        newHover := 0
        if !newHoveredCloseButton {
            for rect in this._windowRects {
                if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                    newHover := rect.actualIndex
                    break
                }
            }
        }

        if newHover != this._mousedOver || newHoveredCloseButton != this._hoveredCloseButton {
            this._mousedOver := newHover
            this._hoveredCloseButton := newHoveredCloseButton
            this.__DrawMenu()
        }

        TrackMouseLeave(hwnd) {
            TME_LEAVE := 0x00000002
            size := A_PtrSize = 8 ? 24 : 16    ; TRACKMOUSEEVENT struct size

            tme := Buffer(size, 0)
            NumPut('UInt', size,          tme, 0)                     ; cbSize
            NumPut('UInt', TME_LEAVE,     tme, 4)                     ; dwFlags
            NumPut('Ptr',  hwnd,          tme, 8)                     ; hwndTrack
            NumPut('UInt', 0,             tme, A_PtrSize = 8 ? 16 : 12)
            return tme
        }
    }

    static __OnLeftClick(wParam, lParam, msg, hwnd) {
        x := lParam & 0xFFFF
        y := lParam >> 16

        ; check if clicking in the search bar
        if this._searchText = this.defaultSearchText {
            if this.searchBackgroundColor != this.searchTextColor {
                rect := this.searchBackgroundRect
                if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.h {
                    this._searchText := ''
                    this.__DrawMenu()
                    return
                }
            }
        }

        ; check if clicking a close button
        for rect in this._closeButtonRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                this._leftClicked := 'close' rect.actualIndex
                return
            }
        }

        ; otherwise check for row clicks
        for rect in this._windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                this._leftClicked := 'activate' rect.actualIndex
                break
            }
        }
    }

    static __OnLeftClickRelease(wParam, lParam, msg, hwnd) {
        x := lParam & 0xFFFF
        y := lParam >> 16

        ; check if clicking a close button
        for rect in this._closeButtonRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                if 'close' . rect.actualIndex = this._leftClicked {
                    this.__CloseWindow(this.menu.windows[rect.actualIndex])
                }
                return
            }
        }

        ; otherwise check for row clicks
        for rect in this._windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                if 'activate' . rect.actualIndex = this._leftClicked {
                    this.CloseMenu()
                    this.__ActivateWindow(rect.window)
                }
                break
            }
        }
    }

    static __OnKeyPress(ih, vk, sc) {
        matches := this._allWindows.Clone()
        key := GetKeyName(Format('vk{:x}sc{:x}', vk, sc))

        switch key {
        case 'Escape':
            if this.escapeAlwaysClose || this._searchText = this.defaultSearchText {
                this.CloseMenu()
                return
            }

            this._searchText := this.defaultSearchText

        case 'Enter':
            this.ActivateWindowAndCloseMenu()
            return

        case 'Backspace':
            if GetKeyState('Control') {
                this._searchText := this.defaultSearchText
            } else if this._searchText != this.defaultSearchText {
                this._searchText := SubStr(this._searchText, 1, -1)
            }

        case 'Up':
            this.SelectPreviousWindow()
            return

        case 'Down':
            this.SelectNextWindow()
            return

        case 'Space':
            this.__AddInputCharacter(' ')

        case 'Delete':
            this.__CloseWindow(this.menu.windows[this._selectedRow])
            return

        default:
            ; get the actual character including shift state
            char := __GetCharFromVK(vk, sc)

            if StrLen(char) = 1 && char != '' {
                this.__AddInputCharacter(char)
            } else {
                ; ignore special keys
                return
            }
        }

        if this._searchText = this.defaultSearchText {
            ; do nothing
        } else if StrLen(this._searchText) = 0 {
            this._searchText := this.defaultSearchText
        } else {
            matches := []
            for window in this._allWindows {
                if InStr(window.name, this._searchText) || InStr(window.title, this._searchText) {
                    matches.Push(window)
                }
            }
        }

        this.menu.windows := matches
        this.__UpdateMenuWhenFiltered()

        __GetCharFromVK(vk, sc) {
            ; get keyboard state
            keyState := Buffer(256, 0)
            DllCall('GetKeyboardState', 'Ptr', keyState)

            ; convert VK to character
            charBuf := Buffer(2, 0)
            result := DllCall('ToUnicode', 'UInt', vk, 'UInt', sc, 'Ptr', keyState, 'Ptr', charBuf, 'Int', 2, 'UInt', 0)

            if result > 0 {
                return StrGet(charBuf, result, 'UTF-16')
            }
            return ''
        }
    }

    static __AddInputCharacter(input) {
        if this._searchText = this.defaultSearchText {
            this._searchText := StrUpper(input)
        } else {
            this._searchText .= StrUpper(input)
        }
    }


    static __OnMouseWheel(wParam, lParam, msg, hwnd) {
        ; get scroll direction
        wheelDelta := (wParam >> 16) & 0xFFFF
        if wheelDelta > 0x7FFF {
            wheelDelta := wheelDelta - 0x10000
        }

        ; update target scroll position
        scrollAmount := 40

        if wheelDelta > 0 {
            this._targetScrollOffset -= scrollAmount
        } else {
            this._targetScrollOffset += scrollAmount
        }

        ; clamp target
        rowWithDivider := this._rowHeight + this._dividerHeight
        totalContentHeight := this.menu.windows.Length * rowWithDivider
        visibleHeight := this.__CalculateTotalHeight() - this._bannerHeight
        maxScrollPixels := Max(0, totalContentHeight - visibleHeight)
        this._targetScrollOffset := Max(0, Min(this._targetScrollOffset, maxScrollPixels))

        ; start animation if not already running
        if !this._scrollTimer {
            this._scrollTimer := ObjBindMethod(this, '__AnimateScroll')
            SetTimer(this._scrollTimer, 8)  ; Increased from 16ms to 8ms (~120 FPS)
        }
    }

    static __AnimateScroll() {
        diff := this._targetScrollOffset - this._scrollOffset

        ; if close enough, snap to target and stop
        if Abs(diff) < 0.5 {
            this._scrollOffset := this._targetScrollOffset
            SetTimer(this._scrollTimer, 0)
            this._scrollTimer := 0
            this.__DrawMenu()
            this.__UpdateHoverFromMouse()
            return
        }

        ; faster easing for snappier feel (increased from 0.2 to 0.3-0.4)
        this._scrollOffset += diff * 0.35
        this.__DrawMenu()
        this.__UpdateHoverFromMouse()
    }

    static __UpdateHoverFromMouse() {
        ; get current mouse position
        CoordMode('Mouse', 'Screen')
        MouseGetPos(&mouseX, &mouseY)

        ; convert screen to client coordinates
        pt := Buffer(8)
        NumPut('Int', mouseX, pt, 0)
        NumPut('Int', mouseY, pt, 4)
        DllCall('ScreenToClient', 'Ptr', this.menu.Hwnd, 'Ptr', pt)

        x := NumGet(pt, 0, 'Int')
        y := NumGet(pt, 4, 'Int')

        ; check which row the mouse is over
        newHover := 0
        for rect in this._windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                newHover := rect.actualIndex
                break
            }
        }

        if newHover != this._mousedOver {
            this._mousedOver := newHover
            this.__DrawMenu()
        }
    }

    static __OnMouseLeave(*) {
        this._mousedOver := 0
        this._mouseLeft := true
        this.__DrawMenu()
    }

    static __ScrollUp() {
        if this._scrollOffset > 0 {
            this._scrollOffset--
            this.__DrawMenu()
        }
    }

    static __Scroll(amount) {
        rowWithDivider := this._rowHeight + this._dividerHeight
        totalContentHeight := this.menu.windows.Length * rowWithDivider
        visibleHeight := this.__CalculateTotalHeight() - this._bannerHeight
        maxScrollPixels := Max(0, totalContentHeight - visibleHeight)

        this._scrollOffset := Max(0, Min(this._scrollOffset + amount, maxScrollPixels))
        this.__DrawMenu()
    }

    static __ActivateWindow(window) {
        if this._selectedRow > 0 && this._selectedRow <= this.menu.windows.Length {
            WinActivate(window)
            return true
        }
    }

    static __ScrollToSelectedRow() {
        rowWithDivider := this._rowHeight + this._dividerHeight
        selectedRowTop := (this._selectedRow - 1) * rowWithDivider
        selectedRowBottom := selectedRowTop + this._rowHeight

        visibleHeight := this.__CalculateTotalHeight() - this._bannerHeight

        ; scroll if selected row is above visible area
        if selectedRowTop < this._scrollOffset {
            this._scrollOffset := selectedRowTop
        }
        ; scroll if selected row is below visible area
        else if selectedRowBottom > this._scrollOffset + visibleHeight {
            this._scrollOffset := selectedRowBottom - visibleHeight
        }
    }


    /**
     * @author iseahound - modernized
     * @source - https://www.autohotkey.com/boards/viewtopic.php?f=83&p=566016#p566016
     * Modified
     * @returns {Array}
     */
    static __AltTabWindows() {
        static  WS_EX_TOOLWINDOW := 0x80,
                GW_OWNER         := 4

        AltTabList := []
        DetectHiddenWindows(false)

        for hwnd in WinGetList() {
            owner := DllCall('GetAncestor', 'Ptr', hwnd, 'UInt', GA_ROOTOWNER := 3, 'Ptr')
            owner := owner || hwnd

            if DllCall('GetLastActivePopup', 'Ptr', owner) = hwnd {
                ex := WinGetExStyle(hwnd)

                if !DllCall('IsWindowVisible', 'Ptr', hwnd) {
                    continue
                }

                if (ex & WS_EX_TOOLWINDOW) {
                    continue
                }

                title := WinGetTitle(hwnd)
                if (title = '') {
                    continue
                }

                AltTabList.Push(hwnd)
            }
        }

        return AltTabList
    }

    static __GDIP_Cleanup() {
        if this._pGraphics {
            Gdip_DeleteGraphics(this._pGraphics)
            this._pGraphics := 0
        }
        if this._hdc {
            SelectObject(this._hdc, this.obm)
            DeleteDC(this._hdc)
            this._hdc := 0
        }
        if this._hBitmap {
            DeleteObject(this._hBitmap)
            this._hBitmap := 0
        }

        ; cleanup cached icons - extract bitmap from object
        for key, iconData in this._iconCache {
            if iconData && IsObject(iconData) && iconData.HasProp('bitmap') && iconData.bitmap {
                Gdip_DisposeImage(iconData.bitmap)
            } else if iconData && !IsObject(iconData) {
                ; handle old cached items that are just pointers
                Gdip_DisposeImage(iconData)
            }
        }
        this._iconCache := Map()
    }


    static __New() {
        this.menu := Gui('+AlwaysOnTop +ToolWindow -SysMenu -Caption +E0x80000')

        this._OnMouseMove        := ObjBindMethod(this, '__OnMouseMove')
        this._OnMouseWheel       := ObjBindMethod(this, '__OnMouseWheel')
        this._OnMouseLeave       := ObjBindMethod(this, '__OnMouseLeave')
        this._OnLeftClick        := ObjBindMethod(this, '__OnLeftClick')
        this._OnLeftClickRelease := ObjBindMethod(this, '__OnLeftClickRelease')

        this._ih := InputHook('L0 V')
        this._ih.KeyOpt('{All}', 'N')
        this._ih.OnKeyDown := ObjBindMethod(this, '__OnKeyPress')

        ; closes task switcher if left click happens outside the menu
        HotIf((*) => TaskSwitcher.isOpen && !TaskSwitcher.hasMouseOver)
        Hotkey('~*LButton', (*) => this.CloseMenu())

        ; GDI+ objects
        this._hBitmap := 0
        this._hdc := 0
        this._pGraphics := 0
        this._pBitmap := 0

        this._sortedWindows := false

        ; cache for icons
        this._iconCache := Map()

        ; store all windows for filtering
        this._allWindows := []

        ; dimensions
        this._maxWidth := 700
        this._bannerHeight := 50
        this._rowHeight := 75
        this._marginX := 12
        this._marginY := 12
        this._iconSize := 32
        this._dividerHeight := 1

        this._scrollOffset := 0
        this._targetScrollOffset := 0
        this._scrollSpeed := 20
        this._maxVisibleRows := 8
        this._hoveredCloseButton := 0
        this._mousedOver := 0
        this._leftClicked := 0
        this._mouseLeft := true
        this._closeButtonRects := []
        this._onWindowActivate := (*) => 0
    }
}