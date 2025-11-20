# Architecture

## Overview
The solution consists of three main components:

1. **VSTO Add-in (`WordOverlayProofreader.Addin`)**:
   - Hosted inside `WINWORD.EXE`.
   - Handles Word Object Model interactions (Text retrieval, Range selection, Coordinate mapping).
   - Sends text to Suggestion Server via WebSocket.
   - Communicates with Overlay (currently via shared state/IPC concept, implemented as separate process for stability).

2. **WPF Overlay (`WordOverlayProofreader.Overlay`)**:
   - Separate process (or topmost window).
   - Transparent, topmost, click-through (mostly).
   - Draws squiggles using WPF vector graphics.
   - Handles user interaction (clicks on squiggles) and shows popovers.

3. **Suggestion Server (`SuggestionServer`)**:
   - Mock backend.
   - WebSocket server.
   - Returns JSON suggestions.

## Data Flow
1. User clicks "Scan" -> Add-in gets text -> Sends to Server.
2. Server processes -> Returns JSON Suggestions.
3. Add-in receives Suggestions -> Calculates Screen Coordinates for each Range -> Sends to Overlay.
4. Overlay draws squiggles.
5. User clicks Squiggle -> Overlay shows Popup.
6. User clicks "Apply" -> Overlay tells Add-in to replace text -> Add-in updates Word -> Re-scans.

## Technologies
- .NET 6 / .NET Framework 4.8 (VSTO requirement).
- VSTO (Visual Studio Tools for Office).
- WPF (Windows Presentation Foundation).
- WebSockets.
