# Windows Utility Appliance Controller- Kill or Restore Windows Updates

## 📸 Screenshot

![Screenshot](https://github.com/Mickle026/Windows10-11_Update_Freeze/blob/master/WindowsUpdateFreeze/WinFreezeUpdates.jpg?raw=true)


A lightweight WinForms tool that lets you turn a Windows machine into a quiet, stable “utility appliance” — and reverse the changes when needed.

Designed for scenarios like:

- Homelabs
- RDP utility machines
- Firmware / vendor tool boxes
- Audio or AI helper nodes
- Compatibility machines alongside Linux
- Systems that must not auto-update or reboot unexpectedly

## ✨ What it does

The application provides simple buttons to:

### 🧊 Freeze (Appliance Mode)
- Disable Windows Update services (`wuauserv`, `usosvc`)
- Apply Windows Update policies to prevent automatic updates and restarts
- Disable common update scheduled tasks:
  - UpdateOrchestrator
  - WindowsUpdate scans
  - InstallService tasks
- Disable common background updaters:
  - Microsoft Edge updater
  - OneDrive updater
  - Firefox background updater
- Reduce update notifications
- Capture a full snapshot so changes can be reversed

### 🔓 Unfreeze (Restore)
- Restore services to their previous state
- Re-enable scheduled tasks as they were
- Restore registry policies
- Return system behaviour to pre-freeze configuration

### ⚡ Power: Always On
- Disable sleep
- Disable hibernate
- Prevent display timeout
- Ideal for headless or remote systems

### 🔄 Power: Restore
- Restore previous power settings from snapshot

### 📊 Status View
- Shows current update service state
- Indicates whether a snapshot exists

## 🧠 How it works

Before making any changes, the app creates a snapshot stored at:


C:\ProgramData\WinUtilityAppliance\snapshot.txt


The snapshot includes:

- Service startup modes
- Whether services were running
- Scheduled task enabled states
- Registry policy values
- Power configuration

This allows the system to be safely restored later.

## 🛡 Safety philosophy

This tool does NOT permanently remove components.

Instead it:

- Disables
- Records state
- Allows reversal

This makes it suitable for lab environments where behaviour needs to be controlled without destructive changes.

## 🧩 Why use this instead of scripts?

Scripts are easy to forget and hard to reverse.

This app provides:

- One-click control
- Visual feedback
- Automatic snapshotting
- Reversible changes
- Repeatable behaviour

## 🏗 Build instructions

### Requirements
- Windows 10 / 11
- .NET 6 or newer SDK
- Visual Studio or `dotnet` CLI

### Build via CLI

dotnet build -c Release

Executable will be in:

bin/Release/net6.0-windows/

Run as Administrator.

🔧 Required package

If building manually, install:

System.ServiceProcess.ServiceController

Example:

dotnet add package System.ServiceProcess.ServiceController
🚀 Usage

Run the app as Administrator

Click Freeze

System enters appliance mode

To restore:

Click Unfreeze

⚠ Important notes

Intended for local or trusted networks

Disabling updates reduces security patching

Recommended for lab / utility systems, not general desktops

Always keep backups or snapshots

🧪 Typical use cases

Linux primary workstation + Windows helper machine

RDP tool server

Offline media tools

Stable automation node

Test environment controller

Long-running processing box

📜 License

Use at your own risk. Provided as-is without warranty.

🤝 Contributions

Feel free to extend:

Add more toggles

Additional service controls

Defender tuning

Telemetry controls

Headless mode presets

Pull requests welcome.

💡 Inspiration

Built for people who want Windows to behave predictably — like firmware — instead of constantly managing itself.
