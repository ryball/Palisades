# Palisade Customization Notes

## Title bar size and startup

1. Right-click a palisade header and choose **Edit**.
2. Use the **Title Bar Size** slider or numeric field to adjust the header height.
3. The title text scales with the header automatically.
4. Use the **Startup** checkbox in the same settings window if you want DeskHaven to launch automatically when you sign in to Windows.

## Fence layout

Palisades now uses a clean, flat layout inside each fence.

- Use separate fences like **Games**, **Tools**, or **Work** for your main categories.
- Drag Windows shortcuts, apps, files, or folders directly into the fence where you want them.
- Drag icons between fences to move them from one fence to another.
- Drag icons back out to the Windows desktop to create desktop shortcuts again, and Palisades now tries to preserve the original shortcut file/icon so the desktop version matches much more closely.
- Reorder shortcuts within the fence by dragging them.
- Right-click a fence header and use **Join as tabs with** to merge two or more fences into one tabbed window.
- Joined fences now show a clearer **tab strip** in the header so the active fence reads more like a real tabbed view.
- Switching tabs now keeps each tab in the **same position** instead of moving the active one to the left.
- You can **drag tabs left or right** inside the header to reorder them manually.
- You can also **drag a tab out of the header** to detach it into its own standalone fence.
- You can also drag one fence onto another to join them automatically as tabs.
- Use **Split tab out** to turn the current tab back into its own standalone fence again.
- Use the **chevron button** in the title bar, or right-click the header and choose **Collapse to title bar**, to shrink a fence down to just its header.
- Use the same chevron button or **Expand fence** to restore the full fence.
- If you use multiple Windows virtual desktops, right-click the fence header and use **Visible on desktops** to choose exactly which named desktops should show that fence.
- Use **Move to desktop** when you want to send the fence to one specific desktop immediately.
- Right-click the fence header and use **Types → Add type...** to create custom types like `Game`, `Tool`, or `Work`.
- Right-click any shortcut icon and use **Type** to assign one of those types, **Add type and assign...**, or **Clear type**.
- Assigned types now show as a small badge under each icon so the category is visible at a glance.
- Right-click any icon and choose **Remove from fence** to remove it from the fence without deleting the original app or file.
- If you remove one by mistake, use **Undo last remove** from the fence menu or press **Ctrl+Z** to restore the last removed icon.
- You can also press **Delete** to remove the currently selected shortcut; modifier keys like `Ctrl` no longer remove icons.
- Right-click the fence header and use **Sort** to order shortcuts by **Name (A to Z)**, **Name (Z to A)**, or **Type**.

## System tray behavior

- Running Palisades now adds an icon to the Windows **system tray**.
- Right-click any fence header and choose **Hide fence** to hide just that one fence.
- Right-click the tray icon to **Show all fences**, **Hide all fences**, **Restore last hidden fences**, or use the **Fences** submenu to **show/hide individual fences** again.
- Double-click the tray icon to show the current fences again.
- On startup, Palisades shows a small tray notification so it is easier to discover where the app is running.

## Persistence

- Header size is stored per palisade.
- Each fence remembers whether it is expanded or collapsed to the title bar.
- Type assignments are stored per shortcut.
- Empty placeholder `No name` fences created by older builds are cleaned up on startup so they do not keep reappearing.

## Installer behavior

- The installer can be run again to **update an existing Palisades install** in place.
- Palisades also exposes an **uninstall** entry in Windows Apps & Features and an **Uninstall Palisades** Start Menu shortcut.
- Uninstall keeps your saved fence data unless you explicitly remove it.
