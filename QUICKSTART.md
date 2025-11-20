# Quick Start Guide

## 🚀 Getting Started in 3 Steps

### Step 1: Build the Project
```bash
cd c:\Users\Admin\Desktop\test-documents
dotnet build WordOverlayProofreader.sln --configuration Debug
```

### Step 2: Run the Applications

#### Option A: Debug from Visual Studio (Recommended)
1. Open `WordOverlayProofreader.sln` in Visual Studio
2. Set `WordOverlayProofreader.Addin` as the startup project
3. Press F5 to start debugging
4. Word will open with the add-in loaded
5. Manually start the Overlay application:
   - Right-click `WordOverlayProofreader.Overlay` → Debug → Start New Instance

#### Option B: Manual Launch
1. **Start Overlay First:**
   ```bash
   cd src/WordOverlayProofreader.Overlay/bin/Debug/net6.0-windows
   ./WordOverlayProofreader.Overlay.exe
   ```
   *(Window will be transparent and stay on top)*

2. **Open Word** and the add-in should load automatically

### Step 3: Start Using

1. Open a Word document or create a new one
2. Type some text with intentional errors: "Ths is a tst document"
3. Click the **"Scan Document"** button in the ribbon
4. Or toggle **"Auto-Scan"** for automatic checking
5. See colored underlines appear
6. Click on underlines to see suggestions
7. Accept suggestions to fix errors!

## 📋 Key Features at a Glance

| Feature | Description |
|---------|-------------|
| **Manual Scan** | Click "Scan Document" button to check current text |
| **Auto-Scan** | Toggle to automatically scan as you type |
| **Color-Coded Errors** | Red (spelling), Blue (grammar), Purple (style), etc. |
| **Interactive Tooltips** | Click underlines to see suggestions |
| **One-Click Fix** | Click suggestion to replace the error |
| **Ignore Option** | Dismiss suggestions you want to keep |

## 🎨 Error Type Colors

- 🔴 Red = Spelling
- 🔵 Blue = Grammar
- 🟣 Purple = Style
- 🟢 Green = Acronym
- 🟠 Orange = Morphology
- 🟤 Brown = Structure
- And more...

## 🔧 Configuration

### Change AI Server URL
Edit `src/WordOverlayProofreader.Addin/SuggestionClient.cs`:
```csharp
private const string ApiUrl = "wss://your-server.com/ws/analyze";
```

### Change Access Tokens
Update the `AccessToken` and `RefreshToken` constants in the same file.

### Adjust Auto-Scan Delay
Edit `src/WordOverlayProofreader.Addin/ThisAddIn.cs`:
```csharp
_scanTimer.Change(2000, Timeout.Infinite); // Change 2000 to your preferred milliseconds
```

## 🐛 Common Issues

### "Add-in not loading in Word"
- Ensure you're running in Debug mode or the add-in is properly installed
- Check Word's Trust Center settings for COM add-ins

### "Overlay not showing underlines"
- Make sure the Overlay application is running
- Check that it's the topmost window
- Try rescanning the document

### "WebSocket connection failed"
- Verify internet connection
- Check if `wss://arabicdemo.abark.tech/ws/analyze` is accessible
- Look for errors in Visual Studio Debug output

### "Clicking Accept does nothing"
- Ensure both Add-in and Overlay are running
- Check that Named Pipes are not blocked
- Look for errors in Debug output

## 📝 Example Workflow

1. **Type:** "Ths is a tst document with misteaks"
2. **Scan:** Click "Scan Document"
3. **See:** Red underlines appear under "Ths", "tst", "misteaks"
4. **Click:** Click on "Ths"
5. **Popup:** Shows suggestion "This"
6. **Accept:** Click the green "This" button
7. **Fixed:** "Ths" is replaced with "This"
8. **Repeat:** Click on "tst" → Accept "test"
9. **Done:** Document is corrected!

## 💡 Pro Tips

- **Enable Auto-Scan** for real-time checking as you type
- **Wait 2 seconds** after typing for auto-scan to trigger
- **Click directly on squiggly lines** for best results
- **Use manual scan** for large documents to control when checking happens
- **Ignore suggestions** you disagree with - they won't come back until next scan

## 🚨 Important Notes

- The overlay must be running separately from the add-in
- Both applications communicate via Named Pipes (local only)
- The AI server requires valid access tokens to work
- First scan may take a few seconds to connect to the server

## 📚 Need More Help?

See `USAGE.md` for detailed documentation or `ARCHITECTURE.md` for technical details.
