# WordOverlayProofreader

A Grammarly-style proofreading overlay for Microsoft Word using VSTO and WPF.

## ✨ Features

✅ **Read Document & Highlight Errors** - Scans text and identifies mistakes via AI server  
✅ **Color-Coded Underlines** - Different colors for different error types (spelling, grammar, style, etc.)  
✅ **Interactive Tooltips** - Click on underlines to see suggestions  
✅ **One-Click Corrections** - Accept suggestions to replace words instantly  
✅ **Auto-Scan Mode** - Automatically checks as you type  
✅ **Manual Scan Button** - Scan on demand via ribbon button

## 🚀 Quick Start

See **[QUICKSTART.md](QUICKSTART.md)** for a 3-step guide to get started immediately!

Or read **[USAGE.md](USAGE.md)** for comprehensive documentation.

## Prerequisites
- Windows 10/11
- Visual Studio 2022 with "Office/SharePoint development" workload
- Word for Windows (Office 365 / 2019)

## Project Structure
- `src/WordOverlayProofreader.Addin`: VSTO Add-in for Word.
- `src/WordOverlayProofreader.Overlay`: WPF Topmost Overlay.
- `src/SuggestionServer`: Mock WebSocket server for suggestions.

## How to Run
1. **Build the Solution**: Open `WordOverlayProofreader.sln` and Build All.
2. **Start Server**: Run `SuggestionServer` project.
3. **Start Overlay**: Run `WordOverlayProofreader.Overlay` project.
4. **Start Add-in**: Set `WordOverlayProofreader.Addin` as startup and press F5 (or install via VSTO installer).
5. **Test**:
   - Open a Word document.
   - Click "Scan Document" in the "Proofreader" ribbon tab.
   - See squiggles and interact with them.

## Quick Run Commands
To run the project from the command line:

1. **Build**:
   ```powershell
   dotnet build WordOverlayProofreader.sln
   ```

2. **Run Server** (in a new terminal):
   ```powershell
   dotnet run --project src/SuggestionServer/SuggestionServer.csproj
   ```

3. **Run Overlay** (in a new terminal):
   ```powershell
   dotnet run --project src/WordOverlayProofreader.Overlay/WordOverlayProofreader.Overlay.csproj
   ```

4. **Run Add-in**:
   - Open `WordOverlayProofreader.sln` in Visual Studio.
   - Set `WordOverlayProofreader.Addin` as the Startup Project.
   - Press **F5**.

## Architecture
See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## License
MIT
