# AHK.ExplorerClassicContextMenu

Note: This class is particularly useful for Windows 11 users who prefer the 
Windows 10 style context menu. It provides a way to restore the classic menu 
with optional notifications for feedback on the process.

/*
	Class: ExplorerClassicContextMenu
	Description: Handles the restoration of the classic (Windows 10 style) context menu in Windows 11
	Usage: 	ExplorerClassicContextMenu(true)  	; Create instance with notifications enabled
			ExplorerClassicContextMenu()      	; Create instance with default notification setting
*/

;! To run this automatically add the script to your library:
<ExplorerClassicContextMenu.ahk>

###Put this somewhere in your main script:
ExplorerClassicContextMenu()

###Or if you want to get notifications:
ExplorerClassicContextMenu(true)
