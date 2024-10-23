/************************************************************************
 * @description Retstore the Windows Explorer Context Menu to the Windows 10 Classic version
 * @author OvercastBTC
 * @date 2024/10/23
 * @version 1.0.0
 ***********************************************************************/
/*
Note: This class is particularly useful for Windows 11 users who prefer the 
Windows 10 style context menu. It provides a way to restore the classic menu 
with optional notifications for feedback on the process.
*/

/**
 * ;! There is a function only script at the bottom. It is commented out.
*/

#Requires AutoHotkey v2.0
#SingleInstance Force

;! To run this automatically un-comment the line below:
; ExplorerClassicContextMenu()

/*
	Class: ExplorerClassicContextMenu
	Description: Handles the restoration of the classic (Windows 10 style) context menu in Windows 11
	Usage: 	ExplorerClassicContextMenu(true)  ; Create instance with notifications enabled
			ExplorerClassicContextMenu()      ; Create instance with default notification setting
*/

Class ExplorerClassicContextMenu {

	; Class property to control notification behavior
	notify := false  ; Default to no notifications

	/*
		Method: __New
		Constructor that initializes the class instance and immediately attempts to restore classic menu
		Parameters:
			notif - Optional boolean to enable/disable notifications (defaults to class property value)
		Example:
			ExplorerClassicContextMenu(true)  ; Enable notifications
			ExplorerClassicContextMenu()      ; Use default (false)
	*/

	__New(notif := this.notify) {

		if notif {
			this.notify := notif
		}
		this.RestoreClassicMenu()
	}

	/*
		Method: InitializeExplorerGroups
		Creates groups of Explorer windows for batch operations
		Used internally before restarting Explorer process
	*/

	InitializeExplorerGroups() {

		GroupAdd("ExplorerGroup", "ahk_class ExploreWClass")    ; Standard Explorer windows
		GroupAdd("ExplorerGroup", "ahk_class CabinetWClass")    ; Explorer windows
		GroupAdd("ExplorerGroup", "ahk_class Progman")          ; Desktop program manager
		GroupAdd("ExplorerGroup", "ahk_class WorkerW")          ; Desktop container
		GroupAdd("ExplorerGroup", "ahk_class #32770")           ; Explorer dialog windows
	}

	/*
		Method: CheckContextMenuState
		Checks if classic context menu is already enabled
		Returns: true if classic menu is enabled, false if modern menu is active
		Also shows notification if this.notify is true
	*/

	CheckContextMenuState() {

		currentValue := unset
		try {
			; Check for registry key that enables classic menu
			currentValue := RegRead("HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32")
			this.notify ? this.trayNotify("Registry key exists.`nWin10 classic Explorer context menu is enabled | Win11 Modern Explorer context menu is disabled",, "T1") : 0
			return true
		} catch {
			return false
		}
	}

	/*
		Method: RestoreClassicMenu
		Main method that handles the entire process of enabling classic context menu
		Steps:
		1. Initialize Explorer window groups
		2. Check current menu state
		3. Create registry key if needed
		4. Restart Explorer to apply changes
	*/

	RestoreClassicMenu() {

		this.InitializeExplorerGroups()

		; Skip if already enabled
		if (this.CheckContextMenuState()) {
			this.notify ? this.trayNotify('No Action Needed' '`n' 'Classic Explorer Context Menu (Windows 10 style) is already enabled.',, 'T1') : 0
			return
		}

		; Create registry key to enable classic menu
		try {
			RegCreateKey("HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32")
			RegWrite("", "REG_SZ", "HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32")
			this.notify ? this.trayNotify("Success!!!'`nClassic Explorer Context Menu has been enabled.`nRestarting Explorer to apply changes...",, "T1") : 0
		} catch as err {
			this.notify ? this.trayNotify("Error:`nFailed to enable Classic Context Menu: " err.Message, "T5") : 0
			return
		}

		; Restart Explorer to apply changes
		try {
			GroupClose("ExplorerGroup")          ; Close all Explorer windows
			Sleep(500)                           ; Allow windows to close
			Run("taskkill.exe /f /im explorer.exe",, "Hide")  ; Force kill Explorer
			Sleep(1000)                          ; Ensure process is terminated
			Run("explorer.exe")                  ; Start new Explorer process
		} catch as err {
			this.notify ? this.trayNotify("Failed to restart Explorer: " err.Message, "Error", "T5") : 0
		}
	}

	/*
		Method: trayNotify
		Enhanced TrayTip function with automatic timeout handling
		Parameters:
			title - Notification title
			message - Optional notification message
			options - TrayTip options (icons etc.) and/or timeout in format "T<seconds>"
			timeout - Optional explicit timeout in milliseconds
		Examples:
			trayNotify("Title", "Message", "T3")      ; Show for 3 seconds
			trayNotify("Title", "Message", 0x10, 5000) ; Show with info icon for 5 seconds
	*/

	trayNotify(title, message := '', options := 0, timeout?) {

		; Extract timeout value if it exists in options string
		if (!IsSet(timeout) && options ~= 'T\d+') {
			RegExMatch(options, "T(\d+)", &match)
			timeout := match[1] * 1000  ; Convert to milliseconds
			options1 := options
			options := RegExReplace(options, "T\d+", "")  ; Remove timeout from options
		}
		else if (IsSet(timeout)) {
			if (Abs(timeout) < 1000) {
			; Handle sub-second timeouts by converting to "T" format
			timeoutSeconds := Abs(timeout) / 1000
			if (options ~= 'T\d+') {
			RegExMatch(options, "T(\d+)", &match)
			if (timeoutSeconds < match[1]) {
				timeout := match[1] * 1000  ; Use the longer timeout
			}
			options := RegExReplace(options, "T\d+", "")
			}
			} 
			else if (options ~= 'T\d+') {
				RegExMatch(options, "T(\d+)", &match)
				if (Abs(timeout/1000) > match[1]) {
				timeout := match[1] * 1000  ; Use the shorter timeout
				}
				options := RegExReplace(options, "T\d+", "")
			}
		}
		TrayTip(title, message, options)
		if (IsSet(timeout)) {
			SetTimer(this.HideTrayTip, -timeout)
		}
	}

	/*
		Method: HideTrayTip
		Helper method to properly hide tray notifications
		Uses multiple approaches to ensure notification is hidden
	*/
	
	HideTrayTip() {
		A_IconHidden := true
		Sleep(500)
		A_IconHidden := false
		TrayTip()
		DllCall("Shell32\Shell_NotifyIconGetRect", "UInt", 0, "Ptr", 0)
		DllCall("User32\UpdateWindow", "Ptr", A_ScriptHwnd)
	}
}

; ; Define Explorer window groups
; InitializeExplorerGroups() {
; 	GroupAdd("ExplorerGroup", "ahk_class ExploreWClass")    ; Explorer windows
; 	GroupAdd("ExplorerGroup", "ahk_class CabinetWClass")    ; Explorer windows
; 	GroupAdd("ExplorerGroup", "ahk_class Progman")          ; Desktop program manager
; 	GroupAdd("ExplorerGroup", "ahk_class WorkerW")          ; Desktop container
; 	GroupAdd("ExplorerGroup", "ahk_class #32770")          ; Explorer dialog windows
; }

; CheckContextMenuState() {
; 	currentValue := unset
; 	try {
; 		currentValue := RegRead("HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32")
; 		; trayNotify("Debug:`nCurrent value: '" currentValue "'`nRegistry key exists - Modern menu is disabled (Classic is enabled)",, "T1")
; 		return true  ; Registry key exists = Classic menu is enabled
; 	} catch {
; 		; trayNotify("Debug:`nCurrent value: " currentValue "`nRegistry key doesn't exist - Modern menu is enabled",, "T1")
; 		return false  ; Registry key doesn't exist = Modern menu is enabled
; 	}
; }

; RestoreClassicMenu() {
; 	; Initialize Explorer groups
; 	InitializeExplorerGroups()

; 	; Check current state
; 	if (CheckContextMenuState()) {
; 		; trayNotify("No Action Needed'`n'Classic Context Menu (Windows 10 style) is already enabled.",, 'T1')
; 		return
; 	}

; 	; Enable classic context menu
; 	try {
; 		RegCreateKey("HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32")
; 		RegWrite("", "REG_SZ", "HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32")
; 		trayNotify("Success!!!'`nClassic Context Menu has been enabled.`nRestarting Explorer to apply changes...",, "T1")
; 	}
; 	catch as err {
; 		trayNotify("Error:`nFailed to enable Classic Context Menu: " err.Message,"T5")
; 		return
; 	}

; 	; Restart Explorer
; 	try {
; 		; Close all Explorer windows in the group
; 		GroupClose("ExplorerGroup")

; 		; Small delay to allow windows to close
; 		Sleep(500)

; 		; Kill Explorer process
; 		Run("taskkill.exe /f /im explorer.exe", , "Hide")

; 		; Small delay to ensure process is fully terminated
; 		Sleep(1000)

; 		; Start new Explorer process
; 		Run("explorer.exe")
; 	}
; 	catch as err {
; 		MsgBox("Failed to restart Explorer: " err.Message, "Error", "T5")
; 	}
; }

; trayNotify(title, message:='', options := 0, timeout?) {
;     ; Extract timeout value if it exists in options string
;     if (!IsSet(timeout) && options ~= 'T\d+') {
;         RegExMatch(options, "T(\d+)", &match)
;         timeout := match[1] * 1000  ; Convert to milliseconds
;         options1 := options
;         options := RegExReplace(options, "T\d+", "")  ; Remove timeout from options
;     }
;     else if (IsSet(timeout)) {
;         if (Abs(timeout) < 1000) {
;             ; Handle sub-second timeouts by converting to "T" format
;             timeoutSeconds := Abs(timeout) / 1000
;             if (options ~= 'T\d+') {
;                 RegExMatch(options, "T(\d+)", &match)
;                 if (timeoutSeconds < match[1]) {
;                     timeout := match[1] * 1000  ; Use the longer timeout
;                 }
;                 options := RegExReplace(options, "T\d+", "")
;             }
;         } else if (options ~= 'T\d+') {
;             RegExMatch(options, "T(\d+)", &match)
;             if (Abs(timeout/1000) > match[1]) {
;                 timeout := match[1] * 1000  ; Use the shorter timeout
;             }
;             options := RegExReplace(options, "T\d+", "")
;         }
;     }

;     ; Show the tray notification
;     TrayTip(title, message, options)

;     ; Handle timeout
;     if (IsSet(timeout)) {
;         SetTimer(HideTrayTip, -timeout)  ; Negative timeout for single execution
;     }

;     HideTrayTip() {
; 		A_IconHidden := true
; 		Sleep(500)
; 		A_IconHidden := false
;         TrayTip()  ; Attempt to hide using normal method
;         ; Force removing the tray icon
;         DllCall("Shell32\Shell_NotifyIconGetRect", "UInt", 0, "Ptr", 0)
;         ; Additional call that helps ensure the notification is hidden
;         DllCall("User32\UpdateWindow", "Ptr", A_ScriptHwnd)
;     }
; }

; Execute the restore function when script runs
; RestoreClassicMenu()

; Optional hotkey to re-run the script
; #c::RestoreClassicMenu()  ; Windows + C to restore classic menu
