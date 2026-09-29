# Installation & Setup

## Requirements

- Windows
- Windows PowerShell or PowerShell 7+
- No external database
- No internet connection required for normal operation

## Installation

1. Save `AlFajarStudio.ps1` into an application folder.
2. Open PowerShell in that folder.
3. Run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\AlFajarStudio.ps1"
```

## First Run

The application automatically creates:

```text
Data\
Data\bookings.json
```

The JSON file starts as an empty booking collection.

## Backup

Close the application and copy:

```text
Data\bookings.json
```

to a safe backup location.

## Updating

Replace `AlFajarStudio.ps1` with a newer version while keeping the existing
`Data\bookings.json` file so saved bookings remain available.

## Security Note

The application uses a local password login and local booking storage.
Protect the Windows account and application folder with appropriate
operating-system security controls.

© 2026 Al Fajar Studio. All rights reserved.
