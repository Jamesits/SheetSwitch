# SheetSwitch

Quickly switch worksheets in a sidebar for Excel.

## Installation

### Sideloading (Windows)

For development use or loading a specific version.

Installation:

1. Set up a SMB share on your PC, for example, `\\localhost\office-addins`
2. Put `manifest.xml` in your SMB share (e.g. `\\localhost\office-addins\sheetswitch.xml`)
3. Excel -> File -> Options -> Trust Center -> Trust Center Settings -> Trusted Add-in Catalogs -> Catalog Url: `\\localhost\office-addins`, click Add catalog, OK
4. Relaunch Excel
5. Excel -> Insert -> My Add-ins -> Shared Folder -> SheetSwitch, click Add

Upgrade metadata:

Open Excel -> Insert -> My Add-ins -> Shared Folder, Click Refresh on the top right, then select SheetSwitch and click Add again.

## Usage

1. Excel -> Home -> SheetSwitch
