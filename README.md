# Palisades

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

Palisades allow you to organize your desktop icons.

![image](https://user-images.githubusercontent.com/2575182/181373105-3ba42faa-7cf2-4a71-8a9d-c3b330e0e860.png)


## Getting Started

Just download the latest installer on the [Releases](https://github.com/Xstoudi/Palisades/releases) page, use it to install Palisades and run the software.

## Features

- Drag and drop your shortcuts in a Palisade.
- Reorder shortcuts in Palisades by drag and dropping them.
- Individually customize the name, colors, and title bar size for each Palisade.
- Keep each fence clean with a simple flat shortcut layout.
- Add custom types from the fence right-click menu and assign them to individual icons.
- Sort shortcuts from the fence right-click menu by name or type.
- Show a Palisades icon in the Windows system tray with quick actions for showing, hiding, restoring, creating, and exiting fences.
- Creating more and deleting existing Palisades.

## Usage
Just drag and drop shortcuts in a Palisade to add it in. Right click on a Palisade header to edit it, then:

- use the **Title Bar Size** control to resize the header,
- right-click a shortcut and choose **Edit group...** or **Move to existing group**,
- drag a shortcut onto another group's header/body to move it there,
- keep the edit window open while selecting shortcuts to change their **Shortcut Group**,
- collapse or expand groups directly from the fence body.

## Techs used

Palisades was made using .NET 6 and WPF. It uses Material Design In XAML for some part of the UI and Sentry to automagically report issues you could encounters.

Palisade is greatly by [Twometer's NoFences](https://github.com/Twometer/NoFences), which was inspired by [Stardock's Fences](https://www.stardock.com/products/fences/). I didn't want to pay 11€ but I also wanted to train on WPF.
