#Requires AutoHotkey v2.0
#SingleInstance

Suspend(true)
#Include ..\..\lib\Gdip_All.ahk

;============================================================================================
; @Auto_Execute
;============================================================================================
pToken := Gdip_Startup()
OnExit((*) => Gdip_Shutdown(pToken))


;============================================================================================
; @Setup
;============================================================================================
; @example of changing options without modifying the class directly
; see @options near beginning of TaskSwitcher class for all options you can modify
TaskSwitcher({
    backgroundColor: 0xFF111111,
    bannerColor: 0xFF00BB00,
    bannerText: 'Tasks',
    alwaysHighlightFirst: true,
    rowHighlightColor: 0xFF555555,
    defaultSearchText: 'Search...',
    searchBackgroundColor: 0xFF333333
})

Suspend(false)

$F1:: {
    TaskSwitcher.OpenMenu()
    KeyWait('F1')
}


;============================================================================================
; @TaskSwitcher
;============================================================================================
class TaskSwitcher {
    ; @options that can be changed here or used as a property name when passing options to TaskSwitcher({option: value})
    ; Note - options that go through Gdip_TextToGraphics require ARGB format as a string (e.g. 'FF00FF00' is green)
    ;       while other color options use 0xARGB as a hex number (e.g. 0xFF00FF00 is green)
    static sortWindows := true
    static defaultTextColor := 0xFFFFFFFF
    static highlightTextColor := 0xFF6995DB
    static rowHighlightColor := 0x30FFFFFF
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


    static isOpen => WinExist('ahk_id' this.menu.Hwnd)
    static isActive => WinActive('ahk_id' this.menu.Hwnd)
    static hasMouseOver => (MouseGetPos(,, &win), win = TaskSwitcher.menu.Hwnd)

    static Call(options := {}) {
        for option, value in options.OwnProps() {
            this.%option% := value
        }
    }

    static __New() {
        this.menu := Gui('+AlwaysOnTop -SysMenu +ToolWindow -Caption +E0x80000')
        this.menu.BackColor := '333333'

        this.menu.Show('NA')
        WinSetAlwaysOnTop(true, 'ahk_id' this.menu.Hwnd)
        this.menu.Hide()

        this.OnMouseMove := ObjBindMethod(this, '__OnMouseMove')
        this.OnLeftClick := ObjBindMethod(this, '__OnLeftClick')
        this.OnKeyPress  := ObjBindMethod(this, '__OnKeyPress')
        this.OnMouseWheel := ObjBindMethod(this, '__OnMouseWheel')

        ; GDI+ objects
        this.hBitmap := 0
        this.hdc := 0
        this.hGraphics := 0
        this.pBitmap := 0

        ; cache for icons
        this.iconCache := Map()

        ; store all windows for filtering
        this.allWindows := []

        ; keeps track of SetWinEventHook instance
        this._eventHook := 0

        ; dimensions
        this.maxWidth := 700
        this.bannerHeight := 50
        this.rowHeight := 75
        this.marginX := 12
        this.marginY := 12
        this.iconSize := 32
        this.dividerHeight := 1

        this.scrollOffset := 0
        this.targetScrollOffset := 0
        this.scrollSpeed := 20
        this.maxVisibleRows := 8
        this.hoveredCloseButton := 0
        this.closeButtonRects := []
    }

    static OpenMenu() {
        if WinExist('ahk_id ' this.menu.Hwnd) {
            return this.CloseMenu()
        }

        ; setup message handlers
        OnMessage(0x200, this.OnMouseMove)
        OnMessage(0x201, this.OnLeftClick)
        OnMessage(0x20A, this.OnMouseWheel)
        OnMessage(0x100, this.OnKeyPress)
        this._eventHook := SetWinEventHook(__OnWinEvent, 0x0003, 0x0003)

        windows := this.__RefreshWindows()
        this.__CreateMenu()
    }

    static CloseMenu() {
        if !WinExist('ahk_id ' this.menu.Hwnd) {
            return
        }

        OnMessage(0x200, this.OnMouseMove, 0)
        OnMessage(0x201, this.OnLeftClick, 0)
        OnMessage(0x100, this.OnKeyPress,  0)
        OnMessage(0x20A, this.OnMouseWheel, 0)
        this._eventHook.UnHook()

        if this.scrollTimer {
            SetTimer(this.scrollTimer, 0)
            this.scrollTimer := 0
        }

        this.menu.Hide()
        this.__GDIP_Cleanup
    }

    static __CreateMenu() {
        this.windowRects := []
        this.selectedRow := 1
        this.searchText := this.defaultSearchText
        this.scrollOffset := 0
        this.targetScrollOffset := 0    ; reset target
        this.scrollTimer := 0           ; reset timer

        totalHeight := this.__CalculateTotalHeight()
        this.__Init_GDIP(totalHeight)

        ; initial draw
        this.__DrawMenu()

        ; setup window
        this.menu.Show('Hide w' this.maxWidth ' h' totalHeight)
        this.menu.GetPos(&winX, &winY)

        ; center the window manually
        MonitorGetWorkArea(, &left, &top, &right, &bottom)
        centerX := left + (right - left - this.maxWidth) / 2
        centerY := top + (bottom - top - totalHeight) / 2

        FrameShadow(this.menu.Hwnd)

        ; use UpdateLayeredWindow instead of picture control
        UpdateLayeredWindow(this.menu.Hwnd, this.hdc, centerX, centerY, this.maxWidth, totalHeight)
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
        WinClose(window.hwnd)
        WinWaitClose(window.hwnd)
        this.__UpdateMenu()
    }

    static __CalculateTotalHeight() {
        totalRows := this.menu.windows.Length

        if totalRows > this.maxVisibleRows {
            ; Show partial row to indicate scrollability
            visibleRows := this.maxVisibleRows - 0.5  ; Show half of the next row
            totalDividers := Floor(visibleRows)
            contentHeight := Integer(visibleRows * this.rowHeight) + (totalDividers * this.dividerHeight)
        } else {
            totalDividers := Max(0, totalRows - 1)
            contentHeight := (totalRows * this.rowHeight) + (totalDividers * this.dividerHeight)
        }

        totalHeight := this.bannerHeight + contentHeight
        return totalHeight
    }

    static __DrawMenu() {
        totalHeight := this.__CalculateTotalHeight()

        ; clear background
        pBrush := Gdip_BrushCreateSolid(this.backgroundColor)
        Gdip_FillRectangle(this.hGraphics, pBrush, 0, 0, this.maxWidth, totalHeight)
        Gdip_DeleteBrush(pBrush)

        ; draw banner
        pBrushBanner := Gdip_BrushCreateSolid(this.bannerColor)
        Gdip_FillRectangle(this.hGraphics, pBrushBanner, 0, 0, this.maxWidth, this.bannerHeight)
        Gdip_DeleteBrush(pBrushBanner)

        ; draw banner text
        options := 'x' this.marginX ' y16 s18 Bold c' this.bannerTextColor
        Gdip_TextToGraphics(this.hGraphics, this.bannerText, options, 'Arial', this.maxWidth - (this.marginX * 2), this.bannerHeight)


        if this.searchBackgroundColor != this.searchTextColor {
            rect := {
                x: this.maxWidth//2 + this.marginX,
                y: 8,
                w: this.maxWidth//2 - (this.marginX * 2) + 4,
                h: 34,
                r: 8
            }
            this.searchBackgroundRect := rect
            pBrushDebug := Gdip_BrushCreateSolid(this.searchBackgroundColor)
            Gdip_FillRoundedRectangle(this.hGraphics, pBrushDebug, rect.x, rect.y, rect.w, rect.h, rect.r)
            Gdip_DeleteBrush(pBrushDebug)
        }

        ; draw input text (right-aligned)
        displayText := this.searchText . Chr(0x200B)
        inputOptions := 'x' (this.maxWidth - 380) ' y16 Right'
        inputOptions .= (this.searchText = this.defaultSearchText)
            ? 's16 Italic c' this.searchTextColor
            : 's18 Bold c' this.bannerTextColor
        Gdip_TextToGraphics(this.hGraphics, displayText, inputOptions, 'Arial', 370, this.bannerHeight)

        ; set clipping - only draw content below banner
        Gdip_SetClipRect(this.hGraphics, 0, this.bannerHeight, this.maxWidth, totalHeight - this.bannerHeight)

        ; draw window rows with pixel offset
        rowWithDivider := this.rowHeight + this.dividerHeight
        this.windowRects := []
        this.closeButtonRects := []  ; Track close button positions

        for index, window in this.menu.windows {
            ; calculate Y position with scroll offset
            rowY := this.bannerHeight + ((index - 1) * rowWithDivider) - this.scrollOffset

            ; skip only if COMPLETELY outside visible area
            if rowY + this.rowHeight < this.bannerHeight || rowY > totalHeight {
                continue
            }

            this.windowRects.Push({
                x: 0,
                y: Max(rowY, this.bannerHeight),
                w: this.maxWidth,
                h: this.rowHeight,
                window: window,
                actualIndex: index
            })

            ; highlight selected row
            if this.selectedRow = index {
                pBrushHover := Gdip_BrushCreateSolid(this.rowHighlightColor)
                Gdip_FillRectangle(this.hGraphics, pBrushHover, 0, rowY, this.maxWidth, this.rowHeight)
                Gdip_DeleteBrush(pBrushHover)
            }

            ; draw icon
            iconX := this.marginX
            iconY := rowY + (this.rowHeight - this.iconSize) / 2
            this.__DrawIcon(window, iconX, iconY)

            ; draw window title
            textColor := (this.selectedRow = index) ? this.highlightTextColor : this.defaultTextColor
            titleX := iconX + this.iconSize + 15
            programOptions := 'x' titleX ' y' (rowY + 8) ' s18 Bold cFF' SubStr(Format('{:06X}', textColor), 3)
            titleOptions := 'x' titleX ' y' (rowY + 28) ' s16 cFF' SubStr(Format('{:06X}', textColor), 3)
            Gdip_TextToGraphics(this.hGraphics, window.processName, programOptions, 'Arial', this.maxWidth - titleX - this.marginX - 40, this.rowHeight)
            Gdip_TextToGraphics(this.hGraphics, window.title, titleOptions, 'Arial', this.maxWidth - titleX - this.marginX - 40, this.rowHeight)

            ; draw close button (X)
            closeButtonSize := 24
            closeButtonX := this.maxWidth - this.marginX - closeButtonSize - 10
            closeButtonY := rowY + (this.rowHeight - closeButtonSize) / 2

            this.closeButtonRects.Push({
                x: closeButtonX,
                y: closeButtonY,
                w: closeButtonSize,
                h: closeButtonSize,
                actualIndex: index
            })

            ; check if mouse is over this close button
            isHoveringCloseButton := (this.hoveredCloseButton = index)

            ; draw close button background (highlight if hovering)
            if isHoveringCloseButton {
                pBrushClose := Gdip_BrushCreateSolid(0x80FF0000)  ; Semi-transparent red
            } else {
                pBrushClose := Gdip_BrushCreateSolid(0x40FFFFFF)  ; Semi-transparent white
            }
            Gdip_FillEllipse(this.hGraphics, pBrushClose, closeButtonX, closeButtonY, closeButtonSize, closeButtonSize)
            Gdip_DeleteBrush(pBrushClose)

            ; draw X
            pPen := Gdip_CreatePen(isHoveringCloseButton ? 0xFFFFFFFF : 0xFFAAAAAA, 2)
            offset := 6
            Gdip_DrawLine(this.hGraphics, pPen, closeButtonX + offset, closeButtonY + offset, closeButtonX + closeButtonSize - offset, closeButtonY + closeButtonSize - offset)
            Gdip_DrawLine(this.hGraphics, pPen, closeButtonX + closeButtonSize - offset, closeButtonY + offset, closeButtonX + offset, closeButtonY + closeButtonSize - offset)
            Gdip_DeletePen(pPen)

            ; draw divider (skip if outside visible area or last item)
            if index < this.menu.windows.Length {
                dividerY := rowY + this.rowHeight
                if dividerY > this.bannerHeight && dividerY < totalHeight {
                    pBrushDiv := Gdip_BrushCreateSolid(this.dividerColor)
                    Gdip_FillRectangle(this.hGraphics, pBrushDiv, this.marginX, dividerY, this.maxWidth - (this.marginX * 2), this.dividerHeight)
                    Gdip_DeleteBrush(pBrushDiv)
                }
            }
        }

        ; reset clipping
        Gdip_ResetClip(this.hGraphics)

        ; get current window position to maintain it
        this.menu.GetPos(&winX, &winY)
        UpdateLayeredWindow(this.menu.Hwnd, this.hdc, winX, winY, this.maxWidth, totalHeight)
    }

    static __RecreateMenuForFiltering() {
        ; decide which row to highlight
        totalRows := this.menu.windows.Length
        if this.alwaysHighlightFirst {
            this.selectedRow := 1
        } else if this.selectedRow > totalRows {
            this.selectedRow := Max(1, totalRows)
        }

        this.scrollOffset := 0

        this.__GDIP_Cleanup()
        totalHeight := this.__CalculateTotalHeight()
        this.__Init_GDIP(totalHeight)
        this.__DrawMenu()

        ; recenter vertically if height changed
        this.menu.GetPos(&winX, &winY)
        MonitorGetWorkArea(, &left, &top, &right, &bottom)
        centerY := top + (bottom - top - totalHeight) / 2

        UpdateLayeredWindow(this.menu.Hwnd, this.hdc, winX, centerY, this.maxWidth, totalHeight)
    }

    static __UpdateMenu() {
        this.__RefreshWindows()
        this.__RecreateMenuForFiltering()
    }

    static __RefreshWindows() {
        list := this.__AltTabWindows()
        windows := []

        for id in list {
            processName := this.__GetProgramName(WinGetProcessPath(id))
            if processName = '' {
                processName := StrSplit(WinGetProcessName(id), '.exe')[1]
            }
            processTitle := WinGetTitle(id)

            windows.Push({
                hwnd:        id,
                title:       processTitle,
                processName: processName,
            })
        }

        if this.sortWindows {
            SortWindows()
        }
        this.allWindows := windows.Clone()  ; store complete list
        this.menu.windows := windows

        SortWindows() {
            i := 2
            while i <= windows.Length {
                temp := windows[i]
                j := i - 1
                while j >= 1 and StrCompare(windows[j].processName, temp.processName) > 0 {
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
        this.hBitmap := CreateDIBSection(this.maxWidth, totalHeight)
        this.hdc := CreateCompatibleDC()
        this.obm := SelectObject(this.hdc, this.hBitmap)
        this.hGraphics := Gdip_GraphicsFromHDC(this.hdc)
        Gdip_SetSmoothingMode(this.hGraphics, 4)
        Gdip_SetTextRenderingHint(this.hGraphics, 3)
    }

    static __GetProgramName(path) {
        size := DllCall("version\GetFileVersionInfoSizeW", "str", path, "uint*", 0, "uint")
        if !size {
            return
        }

        buf := Buffer(size)
        if !DllCall("version\GetFileVersionInfoW", "str", path, "uint", 0, "uint", size, "ptr", buf) {
            return
        }

        Query(val) {
            ptr := 0, len := 0
            if DllCall("version\VerQueryValueW",
                "ptr", buf,
                "str", "\StringFileInfo\040904b0\" val,
                "ptr*", &ptr,
                "uint*", &len)
            {
                return StrGet(ptr, "UTF-16")
            }
        }

        return Query('ProductName')
    }

    static __DrawIcon(window, x, y) {
        hwnd := window.hwnd
        cacheKey := window.processName

        if !this.iconCache.Has(cacheKey) {
            pBitmap := 0

            ; try to extract icon at higher resolution (256x256) for better quality
            try {
                path := WinGetProcessPath(hwnd)

                ; use PrivateExtractIcons to get large icon (256x256)
                hIcon := 0
                DllCall('PrivateExtractIcons', 'Str', path, 'Int', 0, 'Int', 256, 'Int', 256, 'Ptr*', &hIcon, 'Ptr*', 0, 'UInt', 1, 'UInt', 0)

                if hIcon {
                    pBitmap := Gdip_CreateBitmapFromHICON(hIcon)
                    DllCall('DestroyIcon', 'Ptr', hIcon)
                }

                ; if 256 failed, try 128
                if !pBitmap {
                    hIcon := 0
                    DllCall('PrivateExtractIcons', 'Str', path, 'Int', 0, 'Int', 128, 'Int', 128, 'Ptr*', &hIcon, 'Ptr*', 0, 'UInt', 1, 'UInt', 0)

                    if hIcon {
                        pBitmap := Gdip_CreateBitmapFromHICON(hIcon)
                        DllCall('DestroyIcon', 'Ptr', hIcon)
                    }
                }

                ; if still nothing, try 48
                if !pBitmap {
                    hIcon := 0
                    DllCall('PrivateExtractIcons', 'Str', path, 'Int', 0, 'Int', 48, 'Int', 48, 'Ptr*', &hIcon, 'Ptr*', 0, 'UInt', 1, 'UInt', 0)

                    if hIcon {
                        pBitmap := Gdip_CreateBitmapFromHICON(hIcon)
                        DllCall('DestroyIcon', 'Ptr', hIcon)
                    }
                }
            }

            ; try UWP icon if regular icon failed
            if !pBitmap {
                try {
                    uwpPath := this.__GetLargestUWPLogoPath(hwnd)
                    if uwpPath {
                        pBitmap := Gdip_CreateBitmapFromFile(uwpPath)
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

            this.iconCache[cacheKey] := pBitmap ? pBitmap : 0
        }

        pBitmap := this.iconCache[cacheKey]
        if pBitmap && Gdip_GetImageWidth(pBitmap) {
            ; Enable high quality scaling
            Gdip_SetInterpolationMode(this.hGraphics, 7)  ; HighQualityBicubic
            Gdip_DrawImage(this.hGraphics, pBitmap, x, y, this.iconSize, this.iconSize)
            Gdip_SetInterpolationMode(this.hGraphics, 2)  ; Reset to default
        }
    }

    static __GetLargestUWPLogoPath(hwnd) {
        Address := CallbackCreate(EnumChildProc.Bind(WinGetPID(hwnd)), 'Fast', 2)
        DllCall('User32.dll\EnumChildWindows', 'Ptr', hwnd, 'Ptr', Address, 'UInt*', &ChildPID := 0, 'Int')
        CallbackFree(Address)
        return ChildPID && AppHasPackage(ChildPID) ? GetLargestLogoPath(GetDefaultLogoPath(ProcessGetPath(ChildPID))) : ''

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
            return LoopFilePath
        }
    }

    static __OnMouseMove(wParam, lParam, msg, hwnd) {
        static lastX := 0, lastY := 0

        x := lParam & 0xFFFF
        y := lParam >> 16

        pt := Buffer(8)
        NumPut('Short', x, pt, 0)
        NumPut('Short', y, pt, 4)

        DllCall('ClientToScreen', 'Ptr', hwnd, 'Ptr', pt)

        screenX := NumGet(pt, 0, 'Short')
        screenY := NumGet(pt, 4, 'Short')

        ; Prevents row from updating when typing. Re-rendering the graphics causes this function to be called
        if lastX = screenX && lastY = screenY {
            return
        }

        lastX := screenX
        lastY := screenY

        ; check if hovering over a close button first
        newHoveredCloseButton := 0
        for rect in this.closeButtonRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                newHoveredCloseButton := rect.actualIndex
                break
            }
        }

        ; check which row mouse is over
        newHover := this.selectedRow
        for rect in this.windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                newHover := rect.actualIndex
                break
            }
        }

        if newHover != this.selectedRow || newHoveredCloseButton != this.hoveredCloseButton {
            this.selectedRow := newHover
            this.hoveredCloseButton := newHoveredCloseButton
            this.__DrawMenu()
        }
    }

    static __OnLeftClick(wParam, lParam, msg, hwnd) {
        x := lParam & 0xFFFF
        y := lParam >> 16

        ; check if clicking in the search bar
        if this.searchText = this.defaultSearchText {
            if this.searchBackgroundColor != this.searchTextColor {
                rect := this.searchBackgroundRect
                if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.h {
                    this.searchText := ''
                    this.__DrawMenu()
                    return
                }
            }
        }

        ; check if clicking a close button
        for rect in this.closeButtonRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                this.__CloseWindow(this.menu.windows[rect.actualIndex])
                return
            }
        }

        ; otherwise check for row clicks
        for rect in this.windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                this.__ActivateWindow(rect.window)
                break
            }
        }
    }

    static __OnKeyPress(wParam, lParam, msg, hwnd) {
        matches := this.allWindows.Clone()

        vk := wParam
        sc := (lParam >> 16) & 0xFF
        key := GetKeyName(Format("vk{:x}sc{:x}", vk, sc))

        switch key {
        case 'Escape':
            if this.searchText = this.defaultSearchText {
                this.CloseMenu()
                return
            } else {
                this.searchText := this.defaultSearchText
            }

        case 'Enter':
            if matches.Length > 0 && this.selectedRow > 0 {
                this.__ActivateWindow(this.menu.windows[this.selectedRow])
            }
            return

        case 'Backspace':
            if GetKeyState('Control') {
                this.searchText := this.defaultSearchText
            } else if this.searchText != this.defaultSearchText {
                this.searchText := SubStr(this.searchText, 1, -1)
            }

        case 'NumpadUp':
            this.__SelectPreviousRow()
            return

        case 'NumpadDown':
            this.__SelectNextRow()
            return

        case 'Space':
            this.__AddInputCharacter(' ')

        case 'NumpadDel':
            this.__CloseWindow(this.menu.windows[this.selectedRow])
            return

        default:
            ; get the actual character including shift state
            char := __GetCharFromVK(vk, sc)

            if StrLen(key) = 1 && char != '' {
                this.__AddInputCharacter(char)
            } else {
                return      ; ignore special keys
            }
        }


        if StrLen(this.searchText) = 0 {
            this.searchText := this.defaultSearchText
        } else {
            matches := []
            for window in this.allWindows {
                if InStr(window.processName, this.searchText) || InStr(window.title, this.searchText) {
                    matches.Push(window)
                }
            }
        }

        this.menu.windows := matches
        this.__RecreateMenuForFiltering()

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
        if this.searchText = this.defaultSearchText {
            this.searchText := StrUpper(input)
        } else {
            this.searchText .= StrUpper(input)
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
            this.targetScrollOffset -= scrollAmount
        } else {
            this.targetScrollOffset += scrollAmount
        }

        ; clamp target
        rowWithDivider := this.rowHeight + this.dividerHeight
        totalContentHeight := this.menu.windows.Length * rowWithDivider
        visibleHeight := this.__CalculateTotalHeight() - this.bannerHeight
        maxScrollPixels := Max(0, totalContentHeight - visibleHeight)
        this.targetScrollOffset := Max(0, Min(this.targetScrollOffset, maxScrollPixels))

        ; start animation if not already running
        if !this.scrollTimer {
            this.scrollTimer := ObjBindMethod(this, '__AnimateScroll')
            SetTimer(this.scrollTimer, 8)  ; Increased from 16ms to 8ms (~120 FPS)
        }
    }

    static __AnimateScroll() {
        diff := this.targetScrollOffset - this.scrollOffset

        ; if close enough, snap to target and stop
        if Abs(diff) < 0.5 {
            this.scrollOffset := this.targetScrollOffset
            SetTimer(this.scrollTimer, 0)
            this.scrollTimer := 0
            this.__DrawMenu()
            this.__UpdateHoverFromMouse()
            return
        }

        ; faster easing for snappier feel (increased from 0.2 to 0.3-0.4)
        this.scrollOffset += diff * 0.35
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
        newHover := this.selectedRow
        for rect in this.windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                newHover := rect.actualIndex
                break
            }
        }

        if newHover != this.selectedRow {
            this.selectedRow := newHover
            this.__DrawMenu()
        }
    }

    static __ScrollUp() {
        if this.scrollOffset > 0 {
            this.scrollOffset--
            this.__DrawMenu()
        }
    }

    static __Scroll(amount) {
        rowWithDivider := this.rowHeight + this.dividerHeight
        totalContentHeight := this.menu.windows.Length * rowWithDivider
        visibleHeight := this.__CalculateTotalHeight() - this.bannerHeight
        maxScrollPixels := Max(0, totalContentHeight - visibleHeight)

        this.scrollOffset := Max(0, Min(this.scrollOffset + amount, maxScrollPixels))
        this.__DrawMenu()
    }

    static __ActivateWindow(window, *) {
        this.CloseMenu()
        WinActivate(window)
    }

    static __SelectPreviousRow() {
        if this.selectedRow > 1 {
            this.selectedRow -= 1
            this.__ScrollToSelectedRow()
        } else if this.wrapRowSelection {
            this.selectedRow := this.menu.windows.Length
            this.__ScrollToSelectedRow()
        }

        this.__DrawMenu()
    }

    static __SelectNextRow() {
        if this.selectedRow < this.menu.windows.Length {
            this.selectedRow += 1
            this.__ScrollToSelectedRow()
        } else if this.wrapRowSelection {
            this.selectedRow := 1
            this.__ScrollToSelectedRow()
        }

        this.__DrawMenu()
    }

    static __ScrollToSelectedRow() {
        rowWithDivider := this.rowHeight + this.dividerHeight
        selectedRowTop := (this.selectedRow - 1) * rowWithDivider
        selectedRowBottom := selectedRowTop + this.rowHeight

        visibleHeight := this.__CalculateTotalHeight() - this.bannerHeight

        ; scroll if selected row is above visible area
        if selectedRowTop < this.scrollOffset {
            this.scrollOffset := selectedRowTop
        }
        ; scroll if selected row is below visible area
        else if selectedRowBottom > this.scrollOffset + visibleHeight {
            this.scrollOffset := selectedRowBottom - visibleHeight
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
                if (title = "") {
                    continue
                }

                AltTabList.Push(hwnd)
            }
        }

        return AltTabList
    }

    static __GDIP_Cleanup() {
        if this.hGraphics {
            Gdip_DeleteGraphics(this.hGraphics)
            this.hGraphics := 0
        }
        if this.hdc {
            SelectObject(this.hdc, this.obm)
            DeleteDC(this.hdc)
            this.hdc := 0
        }
        if this.hBitmap {
            DeleteObject(this.hBitmap)
            this.hBitmap := 0
        }

        ; cleanup cached icons
        for key, pBitmap in this.iconCache {
            if pBitmap {
                Gdip_DisposeImage(pBitmap)
            }
        }
        this.iconCache := Map()
    }
}

/**
 * @author @plankoe
 * @source https://old.reddit.com/r/AutoHotkey/comments/18o171o/ahkv2_had_its_first_birthday_as_of_yesterday_its/kekm6xy/
 */
; The callback for SetWinEventHook receives 7 parameters when called.
__OnWinEvent(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
    ; if idObject is 0xFFFFFFF7, this event was triggered by cursor.
    ; Since EVENT_OBJECT_LOCATIONCHANGE is one of the specified events, this function will trigger whenever the cursor is moved.
    ; if you don't want to monitor events from cursor, put a Return.
    if idObject = 0xFFFFFFF7 {
        return
    }

    ; example if you only want events triggered by a window:
    if idObject != 0 {
        return
    }

    ; If you don't want to receive events triggered by a child element, make sure idChild is 0.
    if idChild != 0 {
        return
    }

    if hwnd != TaskSwitcher.menu.Hwnd {
        TaskSwitcher.CloseMenu
    }
}

class SetWinEventHook {
   /**
    * @param function - function to call when event is triggered
    * @param minEvent - minimum event to monitor
    * @param maxEvent - max event to monitor
    * @param pid - specify which process to monitor. Use 0 to receive events from all windows.
    * @param flags - flags that specify which events to skip. Default is 0x2 (Prevent AutoHotkey itself from triggering events). Use 0 to allow AutoHotkey to trigger events.
    */
    __New(function, minEvent := 0x00000001, maxEvent := 0x7FFFFFFF, pid := 0, flags := 0x2) {
        this.callback := CallbackCreate(function, "F", 7)
        this.Hook := DllCall("SetWinEventHook"
            , "UInt" , minEvent       ; UINT eventMin
            , "UInt" , maxEvent       ; UINT eventMax
            , "Ptr"  , 0              ; HMODULE hmodWinEventProc
            , "Ptr"  , this.callback  ; WINEVENTPROC lpfnWinEventProc
            , "UInt" , pid            ; DWORD idProcess
            , "UInt" , 0              ; DWORD idThread
            , "UInt" , flags)         ; UINT dwflags, 0x0|0x2 = OutOfContext|SkipOwnProcess
    }

    UnHook() {
        DllCall("UnhookWinEvent", "Ptr", this.Hook)
        CallbackFree(this.callback)
    }
}