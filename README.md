# CopyToRemoteScript

A PowerShell utility that streamlines copying folders to remote locations using `Robocopy`, with an optional right-click menu integration for ease of use.

## Features

- Quickly copy folders to remote destinations via a context menu or interactive prompt.
- Manage multiple target locations using a CSV-based nickname system.
- Exclude common unwanted files (e.g., `Thumbs.db`, `.DS_Store`) automatically.
- Simple interface to add or remove remote targets.
- Supports logging and error handling.

## Getting Started

1. **Clone the repository:**

   ```bash
   git clone https://github.com/paloj/CopyToRemoteScript.git
   ```

2. **Run the script:**

   Right-click `CopyToRemote.ps1` and select **Run with PowerShell**.

3. **Optional - Add to context menu:**

   Follow the prompt in the script to integrate it with the Windows right-click menu.

## Requirements

- Windows PowerShell
- `Robocopy` (built into Windows)

## License

This project is open-source and available under the [MIT License](LICENSE).
