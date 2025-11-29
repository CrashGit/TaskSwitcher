#Requires AutoHotkey v2.0-a
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

/**
 * @Hotkeys
 * Pick your poison
 */
; Simple hotkey
$F1::TaskSwitcher.OpenMenuSorted()

; Hotkeys to replace AltTab behavior
#HotIf !TaskSwitcher.isActive
!Tab:: {
    TaskSwitcher.OpenMenu()
    TaskSwitcher.SelectNextProgram()
}

#HotIf TaskSwitcher.isActive
!Tab::TaskSwitcher.SelectNextProgram()
+!Tab::TaskSwitcher.SelectPreviousProgram()
*Alt up::TaskSwitcher.ActivateProgramAndCloseMenu()

#HotIf


;============================================================================================
; @TaskSwitcher
;============================================================================================
class TaskSwitcher {
    ; @options that can be changed here or used as a property name when passing options to TaskSwitcher({option: value})
    ; Note - options that go through Gdip_TextToGraphics require ARGB format as a string (e.g. 'FF00FF00' is green)
    ;       while other color options use 0xARGB as a hex number (e.g. 0xFF00FF00 is green)
    static _defaultTextColor := 0xFFFFFFFF
    static _highlightTextColor := 0xFF6995DB
    static _rowHighlightColor := 0x30FFFFFF
    static _backgroundColor := 0xFF333333
    static _dividerColor := 0xFFFFFFFF
    static _bannerColor := 0xFF1B56B5
    static _bannerTextColor := 'FFFFFFFF'
    static _bannerText := 'Task Switcher'
    static _wrapRowSelection := true
    static _alwaysHighlightFirst := false
    static _defaultSearchText := ''
    static _searchTextColor := 'ffc8c8c8'
    static _searchBackgroundColor := this._searchTextColor


    static isOpen => WinExist('ahk_id' this.menu.Hwnd)
    static isActive => WinActive('ahk_id' this.menu.Hwnd)
    static hasMouseOver => (MouseGetPos(,, &win), win = TaskSwitcher.menu.Hwnd)

    static Call(options := {}) {
        for option, value in options.OwnProps() {
            this.%'_' . option% := value
        }
    }

    static __New() {
        this.menu := Gui('+AlwaysOnTop -SysMenu +ToolWindow -Caption +E0x80000')

        this.menu.Show('NA')
        WinSetAlwaysOnTop(true, 'ahk_id' this.menu.Hwnd)
        this.menu.Hide()

        this._OnMouseMove := ObjBindMethod(this, '__OnMouseMove')
        this._OnLeftClick := ObjBindMethod(this, '__OnLeftClick')
        this._OnKeyPress  := ObjBindMethod(this, '__OnKeyPress')
        this._OnMouseWheel := ObjBindMethod(this, '__OnMouseWheel')

        ; closes task switcher if left click happens outside the menu
        HotIf((*) => TaskSwitcher.isOpen && !TaskSwitcher.hasMouseOver)
        Hotkey('~*LButton', (*) => this.CloseMenu())
        HotIf()

        ; GDI+ objects
        this._hBitmap := 0
        this._hdc := 0
        this._hGraphics := 0
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
        this._closeButtonRects := []
    }

    ; sorts the programs alphabetically
    static OpenMenuSorted() {
        this.OpenMenu(true)
    }

    ; uses z-order of programs
    static OpenMenu(sortedWindows := false) {
        if WinExist('ahk_id ' this.menu.Hwnd) {
            return this.CloseMenu()
        }

        this._sortedWindows := sortedWindows

        ; setup message handlers
        OnMessage(0x200, this._OnMouseMove)
        OnMessage(0x201, this._OnLeftClick)
        OnMessage(0x20A, this._OnMouseWheel)
        OnMessage(0x100, this._OnKeyPress)

        this.__RefreshWindows()
        this.__CreateMenu()
    }

    static CloseMenu() {
        if !WinExist('ahk_id ' this.menu.Hwnd) {
            return
        }

        OnMessage(0x200, this._OnMouseMove,  0)
        OnMessage(0x201, this._OnLeftClick,  0)
        OnMessage(0x100, this._OnKeyPress,   0)
        OnMessage(0x20A, this._OnMouseWheel, 0)

        if this._scrollTimer {
            SetTimer(this._scrollTimer, 0)
            this._scrollTimer := 0
        }

        this.menu.Hide()
        this.__GDIP_Cleanup
    }

    static ActivateProgramAndCloseMenu() {
        this.SelectProgram()
        this.CloseMenu()
    }

    static __CreateMenu() {
        this._windowRects := []
        this._selectedRow := 1
        this._searchText := this._defaultSearchText
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

        ; use UpdateLayeredWindow instead of picture control
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
        WinClose(window.hwnd)
        WinWaitClose(window.hwnd)

        this.__RefreshWindows()
        this.__ApplySearchFilter()
        this.__RecreateMenuForFiltering()
    }

    static __ApplySearchFilter() {
        if this._searchText != this._defaultSearchText && StrLen(this._searchText) > 0 {
            matches := []
            for win in this._allWindows {
                if InStr(win.processName, this._searchText) || InStr(win.title, this._searchText) {
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
        pBrush := Gdip_BrushCreateSolid(this._backgroundColor)
        Gdip_FillRectangle(this._hGraphics, pBrush, 0, 0, this._maxWidth, totalHeight)
        Gdip_DeleteBrush(pBrush)

        ; draw banner
        pBrushBanner := Gdip_BrushCreateSolid(this._bannerColor)
        Gdip_FillRectangle(this._hGraphics, pBrushBanner, 0, 0, this._maxWidth, this._bannerHeight)
        Gdip_DeleteBrush(pBrushBanner)

        ; draw banner text
        options := 'x' this._marginX ' y16 s18 Bold c' this._bannerTextColor
        Gdip_TextToGraphics(this._hGraphics, this._bannerText, options, 'Arial', this._maxWidth - (this._marginX * 2), this._bannerHeight)

        if this._searchBackgroundColor != this._searchTextColor {
            rect := {
                x: this._maxWidth//2 + this._marginX,
                y: 8,
                w: this._maxWidth//2 - (this._marginX * 2) + 4,
                h: 34,
                r: 8
            }
            this.searchBackgroundRect := rect
            pBrushDebug := Gdip_BrushCreateSolid(this._searchBackgroundColor)
            Gdip_FillRoundedRectangle(this._hGraphics, pBrushDebug, rect.x, rect.y, rect.w, rect.h, rect.r)
            Gdip_DeleteBrush(pBrushDebug)
        }

        ; draw input text (right-aligned)
        displayText := this._searchText . Chr(0x200B)
        inputOptions := 'x' (this._maxWidth - 380) ' y16 Right'
        inputOptions .= (this._searchText = this._defaultSearchText)
            ? 's16 Italic c' this._searchTextColor
            : 's18 Bold c' this._bannerTextColor
        Gdip_TextToGraphics(this._hGraphics, displayText, inputOptions, 'Arial', 370, this._bannerHeight)

        ; set clipping - only draw content below banner
        Gdip_SetClipRect(this._hGraphics, 0, this._bannerHeight, this._maxWidth, totalHeight - this._bannerHeight)

        ; draw window rows with pixel offset
        rowWithDivider := this._rowHeight + this._dividerHeight
        this._windowRects := []
        this._closeButtonRects := []  ; Track close button positions

        for index, window in this.menu.windows {
            ; calculate Y position with scroll offset
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
                pBrushHover := Gdip_BrushCreateSolid(this._rowHighlightColor)
                Gdip_FillRectangle(this._hGraphics, pBrushHover, 0, rowY, this._maxWidth, this._rowHeight)
                Gdip_DeleteBrush(pBrushHover)
            }

            ; draw icon
            iconX := this._marginX
            iconY := rowY + (this._rowHeight - this._iconSize) / 2
            this.__DrawIcon(window, iconX, iconY)

            ; draw window title
            textColor := (this._selectedRow = index) ? this._highlightTextColor : this._defaultTextColor
            titleX := iconX + this._iconSize + 15
            programOptions := 'x' titleX ' y' (rowY + 8) ' s18 Bold cFF' SubStr(Format('{:06X}', textColor), 3)
            titleOptions := 'x' titleX ' y' (rowY + 28) ' s16 cFF' SubStr(Format('{:06X}', textColor), 3)
            Gdip_TextToGraphics(this._hGraphics, window.processName, programOptions, 'Arial', this._maxWidth - titleX - this._marginX - 40, this._rowHeight)
            Gdip_TextToGraphics(this._hGraphics, window.title, titleOptions, 'Arial', this._maxWidth - titleX - this._marginX - 40, this._rowHeight)

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
            Gdip_FillEllipse(this._hGraphics, pBrushClose, closeButtonX, closeButtonY, closeButtonSize, closeButtonSize)
            Gdip_DeleteBrush(pBrushClose)

            ; draw X
            pPen := Gdip_CreatePen(isHoveringCloseButton ? 0xFFFFFFFF : 0xFFAAAAAA, 2)
            offset := 6
            Gdip_DrawLine(this._hGraphics, pPen, closeButtonX + offset, closeButtonY + offset, closeButtonX + closeButtonSize - offset, closeButtonY + closeButtonSize - offset)
            Gdip_DrawLine(this._hGraphics, pPen, closeButtonX + closeButtonSize - offset, closeButtonY + offset, closeButtonX + offset, closeButtonY + closeButtonSize - offset)
            Gdip_DeletePen(pPen)

            ; draw divider (skip if outside visible area or last item)
            if index < this.menu.windows.Length {
                dividerY := rowY + this._rowHeight
                if dividerY > this._bannerHeight && dividerY < totalHeight {
                    pBrushDiv := Gdip_BrushCreateSolid(this._dividerColor)
                    Gdip_FillRectangle(this._hGraphics, pBrushDiv, this._marginX, dividerY, this._maxWidth - (this._marginX * 2), this._dividerHeight)
                    Gdip_DeleteBrush(pBrushDiv)
                }
            }
        }

        ; reset clipping
        Gdip_ResetClip(this._hGraphics)

        ; get current window position to maintain it
        this.menu.GetPos(&winX, &winY)
        UpdateLayeredWindow(this.menu.Hwnd, this._hdc, winX, winY, this._maxWidth, totalHeight)
    }

    static __RecreateMenuForFiltering() {
        ; decide which row to highlight
        totalRows := this.menu.windows.Length
        if this._alwaysHighlightFirst {
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
        this._hBitmap := CreateDIBSection(this._maxWidth, totalHeight)
        this._hdc := CreateCompatibleDC()
        this.obm := SelectObject(this._hdc, this._hBitmap)
        this._hGraphics := Gdip_GraphicsFromHDC(this._hdc)
        Gdip_SetSmoothingMode(this._hGraphics, 4)
        Gdip_SetTextRenderingHint(this._hGraphics, 3)
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

        if !this._iconCache.Has(cacheKey) {
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

            this._iconCache[cacheKey] := pBitmap ? pBitmap : 0
        }

        pBitmap := this._iconCache[cacheKey]
        if pBitmap && Gdip_GetImageWidth(pBitmap) {
            ; Enable high quality scaling
            Gdip_SetInterpolationMode(this._hGraphics, 7)  ; HighQualityBicubic
            Gdip_DrawImage(this._hGraphics, pBitmap, x, y, this._iconSize, this._iconSize)
            Gdip_SetInterpolationMode(this._hGraphics, 2)  ; Reset to default
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
        for rect in this._closeButtonRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                newHoveredCloseButton := rect.actualIndex
                break
            }
        }

        ; check which row mouse is over
        newHover := this._selectedRow
        for rect in this._windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                newHover := rect.actualIndex
                break
            }
        }

        if newHover != this._selectedRow || newHoveredCloseButton != this._hoveredCloseButton {
            this._selectedRow := newHover
            this._hoveredCloseButton := newHoveredCloseButton
            this.__DrawMenu()
        }
    }

    static __OnLeftClick(wParam, lParam, msg, hwnd) {
        x := lParam & 0xFFFF
        y := lParam >> 16

        ; check if clicking in the search bar
        if this._searchText = this._defaultSearchText {
            if this._searchBackgroundColor != this._searchTextColor {
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
                this.__CloseWindow(this.menu.windows[rect.actualIndex])
                return
            }
        }

        ; otherwise check for row clicks
        for rect in this._windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                this.__ActivateWindow(rect.window)
                break
            }
        }
    }

    static __OnKeyPress(wParam, lParam, msg, hwnd) {
        matches := this._allWindows.Clone()

        vk := wParam
        sc := (lParam >> 16) & 0xFF
        key := GetKeyName(Format("vk{:x}sc{:x}", vk, sc))

        switch key {
        case 'Escape':
            if this._searchText = this._defaultSearchText {
                this.CloseMenu()
                return
            } else {
                this._searchText := this._defaultSearchText
            }

        case 'Enter':
            if this._selectedRow > 0 && this._selectedRow <= this.menu.windows.Length {
                this.__ActivateWindow(this.menu.windows[this._selectedRow])
            }
            return

        case 'Backspace':
            if GetKeyState('Control') {
                this._searchText := this._defaultSearchText
            } else if this._searchText != this._defaultSearchText {
                this._searchText := SubStr(this._searchText, 1, -1)
            }

        case 'NumpadUp':
            this.SelectPreviousProgram()
            return

        case 'NumpadDown':
            this.SelectNextProgram()
            return

        case 'Space':
            this.__AddInputCharacter(' ')

        case 'NumpadDel':
            this.__CloseWindow(this.menu.windows[this._selectedRow])
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

        if this._searchText = this._defaultSearchText {
            ; do nothing
        } else if StrLen(this._searchText) = 0 {
            this._searchText := this._defaultSearchText
        } else {
            matches := []
            for window in this._allWindows {
                if InStr(window.processName, this._searchText) || InStr(window.title, this._searchText) {
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
        if this._searchText = this._defaultSearchText {
            this._searchText := StrUpper(input)
        } else {
            this._searchText .= StrUpper(input)
        }
    }

    static SelectProgram() {
        if this._selectedRow > 0 && this._selectedRow <= this.menu.windows.Length {
            this.__ActivateWindow(this.menu.windows[this._selectedRow])
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
        newHover := this._selectedRow
        for rect in this._windowRects {
            if x >= rect.x && x <= rect.x + rect.w && y >= rect.y && y <= rect.y + rect.h {
                newHover := rect.actualIndex
                break
            }
        }

        if newHover != this._selectedRow {
            this._selectedRow := newHover
            this.__DrawMenu()
        }
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

    static __ActivateWindow(window, *) {
        this.CloseMenu()
        WinActivate(window)
    }

    static SelectPreviousProgram() {
        if this._selectedRow > 1 {
            this._selectedRow -= 1
            this.__ScrollToSelectedRow()
        } else if this._wrapRowSelection {
            this._selectedRow := this.menu.windows.Length
            this.__ScrollToSelectedRow()
        }

        this.__DrawMenu()
    }

    static SelectNextProgram() {
        if this._selectedRow < this.menu.windows.Length {
            this._selectedRow += 1
            this.__ScrollToSelectedRow()
        } else if this._wrapRowSelection {
            this._selectedRow := 1
            this.__ScrollToSelectedRow()
        }

        this.__DrawMenu()
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
                if (title = "") {
                    continue
                }

                AltTabList.Push(hwnd)
            }
        }

        return AltTabList
    }

    static __GDIP_Cleanup() {
        if this._hGraphics {
            Gdip_DeleteGraphics(this._hGraphics)
            this._hGraphics := 0
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

        ; cleanup cached icons
        for key, pBitmap in this._iconCache {
            if pBitmap {
                Gdip_DisposeImage(pBitmap)
            }
        }
        this._iconCache := Map()
    }
}