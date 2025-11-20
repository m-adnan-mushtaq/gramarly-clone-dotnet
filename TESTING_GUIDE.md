# Testing Guide - Step by Step

## 🚀 How to Test the Application

Since we removed VSTO infrastructure, follow these steps to test:

### Step 1: Start the Overlay Application
```bash
cd src\WordOverlayProofreader.Overlay\bin\Debug\net8.0-windows
.\WordOverlayProofreader.Overlay.exe
```
**Expected:** A transparent window opens (you won't see it, but it's running)

### Step 2: Open Microsoft Word
- Open Word application
- Create a new document or open an existing one
- Type some text with errors, for example:
  ```
  Ths is a tst document with som misteaks in it.
  ```

### Step 3: Run the Add-in Test Application
```bash
cd src\WordOverlayProofreader.Addin\bin\Debug\net48
.\WordOverlayProofreader.Addin.exe
```

**You should see console output like:**
```
=== Word Overlay Proofreader Test ===
Make sure:
1. Overlay application is running
2. SuggestionServer is running
3. Word is open with a document

✓ Connected to Word
✓ Add-in initialized
[AddIn] Starting up...
[AddIn] Startup complete

Commands:
  s - Scan document
  a - Toggle auto-scan
  q - Quit
```

### Step 4: Scan the Document
- Press **`s`** to scan the document
- Watch the console output to see the flow:

```
[AddIn] Scanning document...
[AddIn] Document text length: 45 chars
[AddIn] Preview: Ths is a tst document with som misteaks...
[SuggestionClient] ScanAsync called with 45 chars
[SuggestionClient] Connecting to wss://arabicdemo.abark.tech/ws/analyze...
[SuggestionClient] Connected to AI server
[SuggestionClient] Sending request: {"text":"Ths is a tst document...
[SuggestionClient] Request sent, waiting for response...
[SuggestionClient] Received JSON: [{"type":"spelling"...
[SuggestionClient] Parsed 4 suggestions
[SuggestionClient] Invoked SuggestionsReceived event
[AddIn] Received 4 suggestions
[AddIn] Sending 4 visuals to overlay...
[AddIn] Data sent to overlay
```

### Step 5: Check Word Document
- Look at your Word document
- You should see **colored squiggly underlines** under the errors:
  - "Ths" - red underline (spelling)
  - "tst" - red underline (spelling)
  - "som" - red underline (spelling)
  - "misteaks" - red underline (spelling)

### Step 6: Click on Underlines
- Click on any colored underline
- A popup should appear showing:
  - Error type (e.g., "Spelling Error")
  - Original text
  - Suggested correction
  - Accept button (green)
  - Ignore button

### Step 7: Accept a Suggestion
- Click the green **Accept** button
- The word should be replaced in your Word document
- The underline should disappear

---

## 🐛 Troubleshooting

### No underlines appearing?

**Check Console Logs:**

1. **If you see "Could not connect to Word"**
   - Make sure Word is actually open
   - Close the test app and Word, then restart Word first

2. **If you see "No active document"**
   - Create or open a document in Word
   - Type some text with errors

3. **If you see "Failed to connect to overlay"**
   - Make sure the Overlay app is running
   - Restart the Overlay application

4. **If you see WebSocket errors**
   - Check your internet connection
   - The AI server at `wss://arabicdemo.abark.tech/ws/analyze` must be accessible
   - Try accessing it in a browser to verify it's online

5. **If connection succeeds but no suggestions**
   - Check the response from the AI server in console logs
   - The server might not have found any errors (try obvious typos)
   - Verify the access tokens are still valid

### Overlay window not showing underlines?

**Debug steps:**
1. Check console: Do you see "[AddIn] Data sent to overlay"?
2. If NO: Restart the Overlay application first
3. If YES: The overlay might be behind Word window - try Alt+Tab

### Clicking underlines does nothing?

1. Make sure you're clicking directly on the squiggly line
2. Check if the Overlay app is still running
3. Try restarting both applications

---

## 💡 Quick Test Commands

After starting the test application:

- **`s`** - Manually scan the document
- **`a`** - Toggle auto-scan ON/OFF (when ON, scans automatically as you type)
- **`q`** - Quit the application

---

## 📝 Sample Test Text

Copy this into Word to test different error types:

```
Ths is a tst document. Ther are misteaks here that need corection.
The sentance has gramr problems to.
```

Expected errors:
- "Ths" → "This" (spelling)
- "tst" → "test" (spelling)
- "Ther" → "There" (spelling)
- "misteakes" → "mistakes" (spelling)
- "corection" → "correction" (spelling)
- "sentance" → "sentence" (spelling)
- "gramr" → "grammar" (spelling)
- "to" → "too" (grammar)

---

## ✅ Success Indicators

You know it's working when you see:
1. ✅ Console shows successful connection to AI server
2. ✅ Console shows "Received N suggestions"
3. ✅ Console shows "Data sent to overlay"
4. ✅ Colored underlines appear in Word
5. ✅ Clicking underlines shows popup
6. ✅ Accepting suggestions replaces words

---

## 🔄 Restart Everything

If things get stuck:

1. Press `q` in the test app console
2. Close Word
3. Close Overlay app (check Task Manager)
4. Start fresh:
   - Start Overlay
   - Open Word
   - Run test app
   - Press `s` to scan
