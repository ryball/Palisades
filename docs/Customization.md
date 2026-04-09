# Palisade Customization Notes

## Title bar size

1. Right-click a palisade header and choose **Edit**.
2. Use the **Title Bar Size** slider or numeric field to adjust the header height.
3. The title text scales with the header automatically.

## Fence layout

Palisades now uses a clean, flat layout inside each fence.

- Use separate fences like **Games**, **Tools**, or **Work** for your main categories.
- Drag shortcuts directly into the fence where you want them.
- Reorder shortcuts within the fence by dragging them.
- Use the **−** button in the title bar, or right-click the header and choose **Collapse to title bar**, to shrink a fence down to just its header.
- Use the same button or **Expand fence** to restore the full fence.
- If you use multiple Windows virtual desktops, right-click the fence header and use **Visible on desktops** to choose exactly which named desktops should show that fence.
- Use **Move to desktop** when you want to send the fence to one specific desktop immediately.
- Right-click the fence header and use **Types → Add type...** to create custom types like `Game`, `Tool`, or `Work`.
- Right-click any shortcut icon and use **Type** to assign one of those types, **Add type and assign...**, or **Clear type**.
- Assigned types now show as a small badge under each icon so the category is visible at a glance.
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
