# DeskHaven (formerly Palisades)

<p align="center">
  <a href="https://github.com/Xstoudi/Palisades/blob/main/LICENSE">
    <img alt="mit" src="https://img.shields.io/github/license/Xstoudi/Palisades?style=for-the-badge"/>
  </a>
  <a href="https://github.com/Xstoudi/Palisades/releases">
    <img alt="mit" src="https://img.shields.io/github/v/release/Xstoudi/Palisades?label=Version&style=for-the-badge"/>
  </a>
  <a href="https://github.com/Xstoudi/Palisades/releases">
    <img alt="mit" src="https://img.shields.io/github/downloads/Xstoudi/Palisades/total?style=for-the-badge"/>
  </a>
</p>

## Introduction

DeskHaven helps you organize your desktop icons with fences, tabs, tray controls, and quick drag-and-drop workflows.

![image](https://user-images.githubusercontent.com/2575182/181373105-3ba42faa-7cf2-4a71-8a9d-c3b330e0e860.png)


## Getting Started

Download the latest installer on the [Releases](https://github.com/Xstoudi/Palisades/releases) page, run it to install Palisades, and launch the app.

### Install / update / uninstall

Palisades now supports a full **install / update / uninstall** flow:

- run the installer to **install** Palisades the first time,
- run the installer again to **update** an existing install in place,
- uninstall Palisades from **Windows Apps & Features** or by using the included **Uninstall Palisades** Start Menu shortcut,
- local fence data is kept unless you explicitly remove it.

## Features

- Drag desktop shortcuts, apps, files, or folders directly into a DeskHaven fence.
- Reorder shortcuts within a fence, move them between fences, or drag them back to the desktop.
- Individually customize the name, colors, title bar size, and whether DeskHaven starts with Windows.
- Keep each fence clean with a simple flat shortcut layout.
- Add custom types from the fence right-click menu and assign them to individual icons.
- Sort shortcuts from the fence right-click menu by name or type.
- Collapse fences, hide them to the tray, and restore them later.
- Join multiple fences into tabs, reorder tabs, and detach tabs back into standalone fences.
- Choose which Windows virtual desktops each fence should appear on.
- Install, update, and uninstall Palisades cleanly.

## Usage
Just drag and drop shortcuts in a Palisade to add them. Right click on a Palisade header to edit it, then:

- use the **Title Bar Size** control to resize the header,
- sort shortcuts by **Name** or **Type**,
- collapse a fence to its title bar,
- hide fences and restore them from the tray,
- join fences into tabs, drag tabs to reorder them, or drag a tab out to detach it,
- drag shortcuts between fences or back to the desktop,
- right-click any icon and choose **Remove from fence** to take it out without deleting the original app or file,
- use **Undo removed icon** or press **Ctrl+Z** to bring back the last icon you removed from that fence.

## Techs used

Palisades was made using .NET 6 and WPF. It uses Material Design In XAML for some part of the UI and Sentry to automagically report issues you could encounters.

Palisade is greatly by [Twometer's NoFences](https://github.com/Twometer/NoFences), which was inspired by [Stardock's Fences](https://www.stardock.com/products/fences/). I didn't want to pay 11€ but I also wanted to train on WPF.
