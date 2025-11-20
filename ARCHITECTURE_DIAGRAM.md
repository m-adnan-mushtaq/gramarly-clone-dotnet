# System Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Microsoft Word                              │
│                                                                     │
│  ┌───────────────────────────────────────────────────────────┐    │
│  │  Ribbon: "Grammar Checker" Tab                            │    │
│  │  ┌──────────────┐  ┌──────────────┐                      │    │
│  │  │ Scan Document│  │  Auto-Scan   │                      │    │
│  │  │   (Button)   │  │  (Toggle)    │                      │    │
│  │  └──────────────┘  └──────────────┘                      │    │
│  └───────────────────────────────────────────────────────────┘    │
│                                                                     │
│  Document Text:                                                    │
│  "Ths is a tst document"                                           │
│   ^^^      ^^^                                                      │
│   (red)    (red)                                                    │
└─────────────────────────────────────────────────────────────────────┘
                            │
                            │ COM Interop
                            │ (VSTO Add-in)
                            ↓
┌─────────────────────────────────────────────────────────────────────┐
│            WordOverlayProofreader.Addin (VSTO)                      │
│                                                                     │
│  Components:                                                        │
│  • ThisAddIn.cs         - Main logic, document scanning            │
│  • SuggestionClient.cs  - WebSocket client for AI server          │
│  • WordCoordinateHelper - Maps text positions to screen coords    │
│  • Ribbon.cs/xml        - UI controls                              │
│                                                                     │
│  Responsibilities:                                                  │
│  1. Extract document text                                          │
│  2. Send to AI server via WebSocket                               │
│  3. Receive error suggestions                                      │
│  4. Map errors to screen coordinates                               │
│  5. Send visual data to Overlay                                    │
│  6. Apply corrections to Word document                             │
└─────────────────────────────────────────────────────────────────────┘
        │                                              ↑
        │ Named Pipe                                   │ Named Pipe
        │ (Visual Data)                                │ (Suggestion ID)
        ↓                                              │
┌─────────────────────────────────────────────────────────────────────┐
│          WordOverlayProofreader.Overlay (WPF)                       │
│                                                                     │
│  ┌───────────────────────────────────────────────────────────┐    │
│  │ Transparent Overlay Window (TopMost)                      │    │
│  │                                                            │    │
│  │  Canvas with Squiggly Lines:                              │    │
│  │  ~~~~ (Red - Spelling)                                     │    │
│  │  ~~~~ (Blue - Grammar)                                     │    │
│  │  ~~~~ (Purple - Style)                                     │    │
│  │                                                            │    │
│  │  ┌────────────────────────────┐                           │    │
│  │  │  Popup on Click            │                           │    │
│  │  │  ┌──────────────────────┐  │                           │    │
│  │  │  │ Spelling Error       │  │                           │    │
│  │  │  │ Original: Ths        │  │                           │    │
│  │  │  │ Suggestion:          │  │                           │    │
│  │  │  │ ┌──────────────────┐ │  │                           │    │
│  │  │  │ │ This (Accept)    │ │  │                           │    │
│  │  │  │ └──────────────────┘ │  │                           │    │
│  │  │  │ [Ignore]             │  │                           │    │
│  │  │  └──────────────────────┘  │                           │    │
│  │  └────────────────────────────┘                           │    │
│  └───────────────────────────────────────────────────────────┘    │
│                                                                     │
│  Components:                                                        │
│  • OverlayWindow.xaml    - UI layout                               │
│  • OverlayWindow.xaml.cs - Event handlers, pipe listener          │
│  • SquiggleRenderer.cs   - Draw colored underlines                │
└─────────────────────────────────────────────────────────────────────┘
        ↑
        │ WebSocket (WSS)
        │
┌─────────────────────────────────────────────────────────────────────┐
│              AI Server                                              │
│              wss://arabicdemo.abark.tech/ws/analyze                 │
│                                                                     │
│  Input:                                                             │
│  {                                                                  │
│    "text": "Ths is a tst document",                               │
│    "requestId": "abc123",                                          │
│    "accessToken": "...",                                           │
│    "refreshToken": "..."                                           │
│  }                                                                  │
│                                                                     │
│  Output:                                                            │
│  [                                                                  │
│    {                                                                │
│      "type": "spelling",                                           │
│      "text": "Ths",                                                │
│      "suggestion": "This",                                         │
│      "from": 0,                                                    │
│      "to": 3,                                                      │
│      "id": "unique-id",                                            │
│      "requestId": "abc123"                                         │
│    },                                                               │
│    ...                                                              │
│  ]                                                                  │
└─────────────────────────────────────────────────────────────────────┘


═══════════════════════════════════════════════════════════════════════
                         DATA FLOW SEQUENCE
═══════════════════════════════════════════════════════════════════════

1. SCANNING PHASE
   ┌──────┐
   │ User │ types or clicks "Scan Document"
   └──┬───┘
      ↓
   ┌──────────────┐
   │  VSTO Addin  │ extracts document.Content.Text
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │  WebSocket   │ sends text to AI server
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │  AI Server   │ analyzes text, finds errors
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │  WebSocket   │ returns error list
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │  VSTO Addin  │ maps errors to screen coordinates
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │  Named Pipe  │ sends visual data
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │ WPF Overlay  │ draws colored squiggly underlines
   └──────────────┘

2. CORRECTION PHASE
   ┌──────┐
   │ User │ clicks on colored underline
   └──┬───┘
      ↓
   ┌──────────────┐
   │ WPF Overlay  │ shows popup with suggestion
   └──────────────┘
      ↓
   ┌──────┐
   │ User │ clicks "Accept" button
   └──┬───┘
      ↓
   ┌──────────────┐
   │ WPF Overlay  │ sends suggestion ID via Named Pipe
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │  VSTO Addin  │ receives ID, finds suggestion
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │ Word Interop │ range.Text = suggestion
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │    Word      │ text is replaced!
   └──────┬───────┘
          ↓
   ┌──────────────┐
   │  VSTO Addin  │ triggers re-scan after 1 second
   └──────────────┘


═══════════════════════════════════════════════════════════════════════
                         ERROR TYPE COLORS
═══════════════════════════════════════════════════════════════════════

  Error Type       Color        RGB           Hex
  ─────────────────────────────────────────────────────────
  Spelling         Red          255,0,0       #FF0000
  Grammar          Blue         0,0,255       #0000FF
  Style            Purple       128,0,128     #800080
  Acronym          Green        0,128,0       #008000
  Morphology       Orange       255,165,0     #FFA500
  Structure        Brown        165,42,42     #A52A2A
  Formatting       Gray         128,128,128   #808080
  Dialect          Teal         0,128,128     #008080
  Quranic          Gold         255,215,0     #FFD700
  Synonyms         Olive        128,128,0     #808000
  Linguistic       Cyan         0,255,255     #00FFFF


═══════════════════════════════════════════════════════════════════════
                      COMMUNICATION PROTOCOLS
═══════════════════════════════════════════════════════════════════════

1. VSTO ↔ AI Server
   Protocol: WebSocket (WSS)
   URL: wss://arabicdemo.abark.tech/ws/analyze
   Format: JSON
   Security: Bearer token in headers

2. VSTO → Overlay
   Protocol: Named Pipe
   Name: WordOverlayProofreaderPipe
   Direction: Out (from VSTO)
   Format: JSON (SuggestionVisual array)

3. Overlay → VSTO
   Protocol: Named Pipe
   Name: WordOverlayAcceptPipe
   Direction: Out (from Overlay)
   Format: Plain text (suggestion ID)

4. VSTO → Word
   Protocol: COM Interop
   API: Microsoft.Office.Interop.Word
   Operations: Range manipulation, text replacement


═══════════════════════════════════════════════════════════════════════
                         THREAD MODEL
═══════════════════════════════════════════════════════════════════════

VSTO Addin:
  • Main Thread: Word UI thread (STA)
  • WebSocket Thread: Async receive loop
  • Named Pipe Thread: Async acceptance listener
  • Timer Thread: Auto-scan trigger

WPF Overlay:
  • UI Thread: WPF Dispatcher thread (STA)
  • Named Pipe Thread: Async pipe server loop
  • Uses Dispatcher.InvokeAsync for UI updates


═══════════════════════════════════════════════════════════════════════
