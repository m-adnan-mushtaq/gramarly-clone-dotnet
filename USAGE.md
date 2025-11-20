# Word Overlay Proofreader - Usage Guide

## Overview
A Grammarly-like VSTO add-in for Microsoft Word that provides real-time grammar and spelling checking with colored underlines and interactive suggestions, powered by an AI server.

## Features Implemented ✅

### 1. **Read Document & Highlight Wrong Words**
- Automatically scans the document content
- Identifies errors through AI server at `wss://arabicdemo.abark.tech/ws/analyze`
- Tracks error positions in the document

### 2. **Draw Colored Underlines**
Different error types are shown with different colors:
- 🔴 **Red** - Spelling errors
- 🔵 **Blue** - Grammar errors  
- 🟣 **Purple** - Style issues
- 🟢 **Green** - Acronym suggestions
- 🟠 **Orange** - Morphology errors
- 🟤 **Brown** - Structure issues
- ⚫ **Gray** - Formatting issues
- 🔶 **Teal** - Dialect issues
- 🟡 **Gold** - Quranic references
- 🫒 **Olive** - Synonyms
- 🔷 **Cyan** - Linguistic issues

### 3. **Show Tooltip with Suggestions**
- Hover over underlined words to see cursor change to hand
- Click on underlined word to see detailed popup with:
  - Error type
  - Original text
  - Suggested correction
  - Ignore button

### 4. **Replace Word on Accept**
- Click on the suggestion button to accept it
- The word is automatically replaced in the Word document
- Underline disappears after replacement
- Document is re-scanned automatically

## How to Use

### Setup

1. **Build the Solution**
   ```bash
   dotnet build WordOverlayProofreader.sln
   ```

2. **Install the Add-in**
   - Open Word
   - The add-in should load automatically if debugging from Visual Studio
   - Or manually install the VSTO add-in

3. **Start the Overlay Application**
   - The overlay window runs as a separate WPF application
   - It should start automatically or launch it manually
   - It runs transparent and TopMost over Word

### Using the Add-in

1. **Open Word** and you'll see a new "Grammar Checker" tab in the ribbon

2. **Two Ways to Scan:**
   - **Manual Scan**: Click the "Scan Document" button
   - **Auto-Scan**: Toggle the "Auto-Scan" button to enable automatic scanning as you type

3. **Review Suggestions:**
   - Colored squiggly underlines appear under errors
   - Hover to see the hand cursor
   - Click on any underline to see the suggestion popup

4. **Apply Corrections:**
   - Click the green suggestion button to accept the correction
   - Or click "Ignore" to dismiss the suggestion
   - The word is replaced automatically in your document

5. **Auto-Scan Mode:**
   - When enabled, the document is scanned automatically:
     - After you stop typing for 2 seconds
     - When you change selection
     - When document content changes

## Architecture

### Components

1. **WordOverlayProofreader.Addin** (VSTO Add-in)
   - Integrates with Microsoft Word
   - Handles document scanning and text replacement
   - Communicates with AI server via WebSocket
   - Sends visual data to overlay via Named Pipes

2. **WordOverlayProofreader.Overlay** (WPF Application)
   - Transparent overlay window
   - Draws colored squiggly underlines
   - Shows interactive tooltips
   - Sends acceptance back to add-in via Named Pipes

3. **SuggestionServer** (Optional Mock Server)
   - For testing without the real AI server
   - Returns sample suggestions

### Communication Flow

```
Word Document → VSTO Add-in → WebSocket → AI Server
                     ↓                         ↓
              Named Pipe                  Suggestions
                     ↓                         ↓
            WPF Overlay ← ← ← ← ← ← ← ← ← ← ← ↓
                     ↓
              User clicks suggestion
                     ↓
              Named Pipe
                     ↓
         VSTO Add-in replaces text in Word
```

## API Integration

The add-in connects to: `wss://arabicdemo.abark.tech/ws/analyze`

**Request Format:**
```json
{
  "text": "document text here",
  "requestId": "abc123",
  "accessToken": "your-token",
  "refreshToken": "your-refresh-token"
}
```

**Response Format:**
```json
[
  {
    "type": "spelling",
    "text": "teh",
    "suggestion": "the",
    "from": 0,
    "to": 3,
    "occurence": 0,
    "id": "unique-id",
    "requestId": "abc123"
  }
]
```

## Troubleshooting

### Underlines not appearing
- Make sure the Overlay application is running
- Check that Named Pipes are working (no firewall blocking)
- Try clicking "Scan Document" manually

### Suggestions not loading
- Verify internet connection
- Check the AI server is accessible at `wss://arabicdemo.abark.tech/ws/analyze`
- Look at Debug output in Visual Studio for WebSocket errors

### Word replacement not working
- Ensure the document is not protected/read-only
- Check Debug output for errors in ApplySuggestion method

### Overlay blocking mouse clicks
- This is a known limitation of transparent windows
- The overlay only captures clicks on the squiggly lines

## Development

### Adding New Error Types

1. Add color in `SquiggleRenderer.cs`:
```csharp
case "yourtype": return Brushes.YourColor;
```

2. The rest is automatic - the type comes from the AI server

### Customizing Underline Style

Edit `SquiggleRenderer.CreateSquiggle()` method to change:
- Wave length
- Wave height  
- Stroke thickness
- Line style (dotted, dashed, etc.)

## Known Limitations

- Overlay window must be running separately
- Complex document layouts may have coordinate mapping issues
- Large documents may take longer to scan
- Named Pipes work on local machine only

## Future Enhancements

- [ ] Integrated overlay launcher
- [ ] Context menu integration
- [ ] Bulk accept/ignore all
- [ ] Custom dictionary
- [ ] Offline mode
- [ ] Multi-language support
