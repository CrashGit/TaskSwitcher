#Include task switcher.ahk
Suspend(true)


/**
 * Personal note:
 * I like to enact Suspend when loading auto-execute stuff and disable it after. I do this in most of my scripts.
 * That way if a hotkey is triggered too quickly after starting/reloading a script, the hotkey doesn't try to activate something before it's initialized completely and throw an error.
 */


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
    searchBackgroundColor: 0xFF333333,
    highlightTextColor: 0xFF00FF33,
    mouseHighlightTextColor: 0xFF999999,
    escapeAlwaysClose: true
})

TaskSwitcher.OnWindowActivate((window) {
    list := ''
    for prop, value in window.OwnProps() {
        list .= Format('Name: {} - Value: {}`n', prop, value)
    }

    ToolTip(list)
    SetTimer(ToolTip, -3000)
})

Suspend(false)

/**
 * @Hotkeys
 * Pick your poison
 */
; Simple toggle hotkey
$F1::TaskSwitcher.ToggleMenu()

; Simple open and close hotkeys
$F2::TaskSwitcher.OpenMenu()
$F3::TaskSwitcher.CloseMenu()

; Left and right control keys pressed together
$<^RCtrl::TaskSwitcher.ToggleMenuSorted()
$>^LCtrl::TaskSwitcher.ToggleMenuSorted()

; Hotkey setup to replace AltTab behavior
TaskSwitcher.AltTabReplacement()

; Toggles the AltTabReplacement hotkeys created from above. 'On' and 'Off' are also both valid parameters.
$F4::TaskSwitcher.AltTabReplacement('Toggle')