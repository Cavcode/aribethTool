# Aribeth - Tlk and 2da Editor

Desktop editor for Neverwinter Nights 2DA and TLK files with a dark UI and integrated tooling.

## Features
- 2DA editor with add/remove columns, copy/cut/paste rows, and regex search/replace
- TLK editor with JSON round-trip via `nwn_tlk.exe`
- Undo/redo for 2DA and TLK edits
- Hotkeys for common actions
- Built-in tools for 2DA normalize, merge, and issue checks via `nwn-2da.exe`
- Dark theme with custom title bar

## Requirements
- Windows 10/11
- .NET 8 SDK
- External tools in `plugins/`:
  - `nwn_tlk.exe`
  - `nwn-2da.exe`

## Build
Open the solution in Visual Studio or run:
```
dotnet build
```

## Run
```
dotnet run
```

## Usage Notes
- Opening a `.2da` file loads it in the 2DA editor tab.
- Opening a `.tlk` file loads it in the TLK editor tab.
- The TLK editor converts the binary TLK to JSON, edits entries, then converts back on save.
- When no file is open, the center image is shown on the 2DA and TLK tabs.

