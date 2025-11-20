# Testing the Updated App

## What Changed

### 1. Fixed Coordinate Mapping
- `WordCoordinateHelper.GetScreenRect()` now properly converts Word document coordinates to screen coordinates
- Added window position offset (previously returned Empty rectangles)
- Added console logging: `[Coord] Window: (X,Y), Relative: (x,y,w,h), Screen: (screenX,screenY)`
- Default dimensions of 50x20 if Word returns 0

### 2. New Static Suggestion Widget
- Created `SuggestionWidget.xaml` and `SuggestionWidget.xaml.cs`
- Displays at bottom-right corner of screen
- Shows one suggestion at a time with:
  - Error type (colored by type: Spelling=Red, Grammar=Blue, etc.)
  - Original text
  - Suggested correction
  - Counter (e.g., "1 of 12")
  - Previous/Next buttons for navigation
  - Accept/Dismiss buttons
- Integrated into `App.xaml.cs` and `OverlayWindow.xaml.cs`

## How to Test

### Step 1: Clean Restart
```bash
# Kill all running processes
taskkill /F /IM WINWORD.EXE 2>nul
taskkill /F /IM WordOverlayProofreader.Overlay.exe 2>nul
taskkill /F /IM WordOverlayProofreader.Addin.exe 2>nul
```

### Step 2: Start the Overlay
```bash
cd /c/Users/Admin/Desktop/test-documents
./src/WordOverlayProofreader.Overlay/bin/Debug/net8.0-windows/WordOverlayProofreader.Overlay.exe
```

You should see:
```
========================================
  Overlay Application Starting
========================================
[Overlay] Initializing window...
[Overlay] Window initialized successfully
[Overlay] Starting pipe server...
[Overlay] Pipe server thread started
[Overlay] Listening on pipe: WordOverlayProofreaderPipe
[Overlay] Waiting for connection...
[App] Creating SuggestionWidget...
[App] SuggestionWidget created
```

### Step 3: Start Word and the Add-in
```bash
cd /c/Users/Admin/Desktop/test-documents
./src/WordOverlayProofreader.Addin/bin/Debug/net48/WordOverlayProofreader.Addin.exe
```

Press `s` to scan document (or type Arabic text first if document is empty)

### Step 4: Look for the NEW Coordinate Logs

In the console, you should now see:
```
[Coord] Window: (100,200), Relative: (15,30,45,18), Screen: (115,230)
✓ Valid suggestion rect: X=115, Y=230, Width=45, Height=18
```

**Instead of:**
```
[Coord] Rect: Empty
✗ Rect is empty or invalid - SKIPPED
```

### Step 5: Check for Widget

Look at the **bottom-right corner** of your screen for a white rounded rectangle with:
- Red text showing error type
- Original text in red box
- Suggestion in green box
- Accept (green) and Dismiss (red) buttons

### Step 6: Test Features

1. **Inline Underlines**: You should now see colored squiggly lines in Word document at the error locations
2. **Widget Navigation**: Click Previous/Next buttons to cycle through suggestions
3. **Accept Suggestion**: Click "✓ Accept" - word should be replaced in document
4. **Dismiss**: Click "✗ Dismiss" - moves to next suggestion without changing anything

## Expected Console Output

### From Add-in:
```
[AddIn] Application assigned successfully
[AddIn] Starting WebSocket connection to AI server...
[AddIn] Connected to AI server
[AddIn] Scanning document...
[AddIn] Suggestions received: 12 suggestions
[AddIn] Suggestion 1: خطاء → خطأ (Spelling) at position 5, ID: 550e8400-e29b-41d4-a716-446655440001
[Coord] Window: (100,200), Relative: (15,30,45,18), Screen: (115,230)
✓ Valid suggestion rect: X=115, Y=230, Width=45, Height=18
[AddIn] Sending 726 bytes to overlay
[AddIn] Connected to overlay pipe
[AddIn] Data sent to overlay
```

### From Overlay:
```
[Overlay] Client connected!
[Overlay] Received 726 bytes
[Overlay] Parsed 12 suggestions
[Overlay] UpdateSuggestions called with 12 items
[Overlay] Adding squiggle at {X=115,Y=230,Width=45,Height=18} for 'خطاء' -> 'خطأ'
[Overlay] Added 12 squiggles to canvas
[Overlay] Updating widget with 12 suggestions
```

## Troubleshooting

### Still seeing "Rect: Empty"?
- Word might not be focused or document might be empty
- Try clicking in the Word document before scanning
- Add some Arabic text if document is empty

### Widget not appearing?
- Check if overlay exe is running
- Look in all corners of screen (should be bottom-right)
- Check console for widget creation messages

### Underlines not visible?
- Make sure Word window is not covering the overlay
- Try minimizing/restoring Word
- Check that coordinates are actually being calculated (not Empty)

## Testing Checklist

- [ ] Overlay starts and shows console
- [ ] Add-in connects to AI server
- [ ] Scanning returns suggestions (12 in test document)
- [ ] Console shows `[Coord] Window:` logs with actual numbers
- [ ] Console shows `✓ Valid suggestion rect:` instead of `✗ SKIPPED`
- [ ] Widget appears at bottom-right
- [ ] Widget shows error type, original, and suggestion
- [ ] Squiggly underlines appear in Word document (colored by type)
- [ ] Clicking Accept replaces the word in document
- [ ] Previous/Next buttons navigate suggestions
- [ ] Dismiss button moves to next suggestion

## Next Steps

If everything works:
1. Test with different types of errors (Grammar, Style, etc.) to verify color coding
2. Test with longer documents with many suggestions
3. Test word replacement with various Arabic text
4. Consider adding hover tooltips for inline squiggles

If there are still issues:
1. Check the console logs for error messages
2. Verify coordinate calculations are actually happening
3. Check that Word window is visible and active
4. Try with a new empty document and type test text
