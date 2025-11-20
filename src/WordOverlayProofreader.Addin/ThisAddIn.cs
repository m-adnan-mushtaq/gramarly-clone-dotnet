using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordOverlayProofreader.Addin
{
    public partial class ThisAddIn
    {
        public Word.Application Application { get; set; }
        
        private SuggestionClient _client;
        private WordCoordinateHelper _coordHelper;
        private System.Threading.Timer _scanTimer;
        private DateTime _lastScanTime = DateTime.MinValue;
        private List<Suggestion> _currentSuggestions = new List<Suggestion>();
        private bool _autoScanEnabled = true; // Enabled by default as requested

        public void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _client = new SuggestionClient();
            _client.SuggestionsReceived += OnSuggestionsReceived;
            _coordHelper = new WordCoordinateHelper(this.Application);
            
            // Hook up events
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
            this.Application.DocumentChange += Application_DocumentChange;
            
            // Start timer for periodic scanning (every 3.5 seconds after changes)
            _scanTimer = new System.Threading.Timer(OnScanTimerTick, null, Timeout.Infinite, Timeout.Infinite);
            
            // Listen for suggestion acceptance from overlay
            Task.Run(() => ListenForAcceptance());
            
            // Initial scan
            if (_autoScanEnabled)
            {
                // Delay initial scan slightly to let Word load
                _scanTimer.Change(2000, Timeout.Infinite);
            }
        }

        private void OnSuggestionsReceived(object sender, List<Suggestion> suggestions)
        {
            try
            {
                Console.WriteLine($"\n=== SUGGESTIONS RECEIVED ===");
                Console.WriteLine($"Total: {suggestions?.Count ?? 0} suggestions\n");
                
                _currentSuggestions = suggestions;
                var visuals = new List<SuggestionVisual>();
                
                var doc = this.Application.ActiveDocument;
                if (doc == null)
                {
                    Console.WriteLine("[AddIn] ERROR: No active document");
                    return;
                }
                
                int index = 1;
                foreach (var s in suggestions)
                {
                    Console.WriteLine($"\n[{index}] Type: {s.type}");
                    Console.WriteLine($"    Text: '{s.text}'");
                    Console.WriteLine($"    Suggestion: '{s.suggestion}'");
                    Console.WriteLine($"    Position: {s.from} to {s.to}");
                    Console.WriteLine($"    ID: {s.id}");
                    
                    // Map range to rect
                    // Word ranges are 1-based, but server returns 0-based indices
                    try
                    {
                        var range = doc.Range(s.from, s.to);
                        var rect = _coordHelper.GetScreenRect(range);
                        
                        Console.WriteLine($"    Rect: {rect}");
                        
                        if (rect != System.Windows.Rect.Empty && rect.Width > 0 && rect.Height > 0)
                        {
                            visuals.Add(new SuggestionVisual
                            {
                                type = s.type,
                                suggestion = s.suggestion,
                                Rect = rect,
                                OriginalText = s.text,
                                id = s.id,
                                from = s.from,
                                to = s.to
                            });
                            Console.WriteLine($"    ✓ Added to visuals");
                        }
                        else
                        {
                            Console.WriteLine($"    ✗ Rect is empty or invalid - SKIPPED");
                        }
                    }
                    catch (Exception rangeEx)
                    {
                        Console.WriteLine($"    ✗ ERROR mapping range: {rangeEx.Message}");
                        System.Diagnostics.Debug.WriteLine($"Error mapping range {s.from}-{s.to}: {rangeEx.Message}");
                    }
                    
                    index++;
                }
                
                Console.WriteLine($"\n=== SUMMARY ===");
                Console.WriteLine($"Total suggestions: {suggestions.Count}");
                Console.WriteLine($"Valid visuals: {visuals.Count}");
                Console.WriteLine($"===============\n");

                SendToOverlay(visuals);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing suggestions: {ex.Message}");
            }
        }

        private void SendToOverlay(List<SuggestionVisual> visuals)
        {
            Task.Run(async () =>
            {
                try
                {
                    Console.WriteLine($"[AddIn] Attempting to connect to overlay pipe...");
                    using (var client = new System.IO.Pipes.NamedPipeClientStream(".", "WordOverlayProofreaderPipe", System.IO.Pipes.PipeDirection.Out))
                    {
                        await client.ConnectAsync(2000);
                        Console.WriteLine($"[AddIn] Connected to overlay pipe");
                        using (var writer = new System.IO.StreamWriter(client))
                        {
                            writer.AutoFlush = true;
                            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                            var json = serializer.Serialize(visuals);
                            Console.WriteLine($"[AddIn] Sending {json.Length} bytes to overlay");
                            await writer.WriteAsync(json);
                            await writer.FlushAsync();
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to connect to overlay: {ex.Message}");
                }
            });
        }

        public void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            _scanTimer?.Dispose();
            _client?.Dispose();
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            // Trigger re-scan after a delay if auto-scan is enabled
            if (_autoScanEnabled)
            {
                // Debounce 3.5 seconds
                _scanTimer.Change(3500, Timeout.Infinite);
            }
        }

        private void Application_DocumentChange()
        {
            // Document content changed, trigger scan if auto-scan is enabled
            if (_autoScanEnabled)
            {
                // Debounce 3.5 seconds
                _scanTimer.Change(3500, Timeout.Infinite);
            }
        }
        
        private void OnScanTimerTick(object state)
        {
            try
            {
                // Prevent too frequent scans
                if ((DateTime.Now - _lastScanTime).TotalSeconds < 1)
                    return;
                    
                _lastScanTime = DateTime.Now;
                ScanDocument();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Timer scan error: {ex.Message}");
            }
        }

        public async void ScanDocument()
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc == null) return;
                
                var text = doc.Content.Text;
                
                // Send to server
                await _client.ScanAsync(text);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ScanDocument error: {ex.Message}");
            }
        }
        
        public void ApplySuggestion(string suggestionId)
        {
            try
            {
                Console.WriteLine($"[AddIn] ApplySuggestion called with ID: {suggestionId}");
                Console.WriteLine($"[AddIn] Current suggestions count: {_currentSuggestions.Count}");
                
                var suggestion = _currentSuggestions.FirstOrDefault(s => s.id == suggestionId);
                if (suggestion == null)
                {
                    Console.WriteLine($"[AddIn] ERROR: Suggestion ID not found: {suggestionId}");
                    return;
                }
                
                Console.WriteLine($"[AddIn] Found suggestion: '{suggestion.text}' -> '{suggestion.suggestion}' at position {suggestion.from}-{suggestion.to}");
                
                var doc = this.Application.ActiveDocument;
                if (doc == null)
                {
                    Console.WriteLine($"[AddIn] ERROR: No active document");
                    return;
                }
                
                // Find and replace the text
                Console.WriteLine($"[AddIn] Replacing text in document...");
                var range = doc.Range(suggestion.from, suggestion.to);
                Console.WriteLine($"[AddIn] Original range text: '{range.Text}'");
                range.Text = suggestion.suggestion;
                Console.WriteLine($"[AddIn] ✓ Text replaced successfully!");
                
                // Remove from current suggestions
                _currentSuggestions.Remove(suggestion);
                Console.WriteLine($"[AddIn] Removed suggestion from list. Remaining: {_currentSuggestions.Count}");
                
                // Re-scan document after a short delay
                _scanTimer.Change(1000, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[AddIn] ApplySuggestion ERROR: {ex.Message}");
                Console.WriteLine($"[AddIn] Stack trace: {ex.StackTrace}");
                System.Diagnostics.Debug.WriteLine($"ApplySuggestion error: {ex.Message}");
            }
        }
        
        public void SetAutoScan(bool enabled)
        {
            _autoScanEnabled = enabled;
            if (enabled)
            {
                // Scan immediately when auto-scan is enabled
                ScanDocument();
            }
            else
            {
                // Cancel any pending scans
                _scanTimer.Change(Timeout.Infinite, Timeout.Infinite);
            }
        }
        
        private async Task ListenForAcceptance()
        {
            while (true)
            {
                try
                {
                    using (var server = new System.IO.Pipes.NamedPipeServerStream("WordOverlayAcceptPipe", System.IO.Pipes.PipeDirection.In))
                    {
                        Console.WriteLine("[AddIn] Waiting for acceptance pipe connection...");
                        await server.WaitForConnectionAsync();
                        Console.WriteLine("[AddIn] Acceptance pipe connected!");
                        using (var reader = new System.IO.StreamReader(server))
                        {
                            var suggestionId = await reader.ReadToEndAsync();
                            Console.WriteLine($"[AddIn] Received suggestion ID: '{suggestionId}'");
                            if (!string.IsNullOrEmpty(suggestionId))
                            {
                                ApplySuggestion(suggestionId.Trim());
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ListenForAcceptance error: {ex.Message}");
                    await Task.Delay(1000);
                }
            }
        }

        // Singleton instance
        private static ThisAddIn _instance;
        public static ThisAddIn Instance => _instance ?? (_instance = new ThisAddIn());
    }

    public class SuggestionVisual
    {
        public string type { get; set; }
        public string suggestion { get; set; }
        public System.Windows.Rect Rect { get; set; }
        public string OriginalText { get; set; }
        public string id { get; set; }
        public int from { get; set; }
        public int to { get; set; }
    }
}
