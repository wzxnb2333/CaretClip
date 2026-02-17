# CaretClip

A Win11-style clipboard history tool for Windows (WPF).

## Features

- Global hotkey: `Alt + Shift + V`
- Popup clipboard history panel
- Search, pin, delete, and copy history entries
- Edit text clipboard items in a dedicated editor window
- Retention period settings (1 day / 7 days / 30 days / forever)
- Startup toggle and tray icon support
- Light / Dark / Follow system theme

## Build

```powershell
dotnet build CopyBoard/CopyBoard.csproj -c Release
```

## Run

```powershell
dotnet run --project CopyBoard/CopyBoard.csproj
```

## Publish (win-x64)

```powershell
dotnet publish CopyBoard/CopyBoard.csproj -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true -o artifacts/CaretClip-win-x64
```

## Project Structure

- `CopyBoard/` - WPF application source code
- `artifacts/` - local publish outputs

## License

MIT. See `LICENSE`.

