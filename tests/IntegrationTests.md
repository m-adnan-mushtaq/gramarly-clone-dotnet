# Integration Tests

## Manual Integration Test

1. **Start the Server**
   - Run `SuggestionServer.exe`
   - Verify output: `Suggestion Server started on ws://localhost:8080/`

2. **Start the Overlay**
   - Run `WordOverlayProofreader.Overlay.exe`
   - It should run silently (transparent window).

3. **Start Word with Add-in**
   - Open Word.
   - Check the "Proofreader" tab in Ribbon.
   - Click "Scan Document".

4. **Verify Interaction**
   - Check `SuggestionServer` console for "Client connected" and "Received: ...".
   - Check Word overlay for squiggles (red/blue lines).
   - Click a squiggle to see the popup.
   - Click a suggestion to apply it.

## Automated Integration Test (Future)
- Use `Appium` or `WinAppDriver` to drive Word and the Overlay.
- Verify elements exist on screen.
