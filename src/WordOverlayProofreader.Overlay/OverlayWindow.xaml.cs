using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace WordOverlayProofreader.Overlay
{
    public partial class OverlayWindow : Window
    {
        private const string PipeName = "WordOverlayProofreaderPipe";

        public OverlayWindow()
        {
            Console.WriteLine("[Overlay] Initializing window...");
            try
            {
                InitializeComponent();
                Console.WriteLine("[Overlay] Window initialized successfully");
                Console.WriteLine("[Overlay] Starting pipe server...");
                Task.Run(() => StartPipeServer());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Overlay] ERROR in constructor: {ex.Message}");
                Console.WriteLine($"[Overlay] Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        private async Task StartPipeServer()
        {
            Console.WriteLine("[Overlay] Pipe server thread started");
            Console.WriteLine($"[Overlay] Listening on pipe: {PipeName}");
            
            while (true)
            {
                try
                {
                    Console.WriteLine("[Overlay] Waiting for connection...");
                    using (var server = new System.IO.Pipes.NamedPipeServerStream(PipeName, System.IO.Pipes.PipeDirection.In))
                    {
                        await server.WaitForConnectionAsync();
                        Console.WriteLine("[Overlay] Client connected!");
                        using (var reader = new System.IO.StreamReader(server))
                        {
                            var json = await reader.ReadToEndAsync();
                            Console.WriteLine($"[Overlay] Received {json?.Length ?? 0} bytes");
                            if (!string.IsNullOrEmpty(json))
                            {
                                Console.WriteLine($"[Overlay] JSON data: {json.Substring(0, Math.Min(200, json.Length))}...");
                                await Dispatcher.InvokeAsync(() =>
                                {
                                    try 
                                    {
                                        var suggestions = System.Text.Json.JsonSerializer.Deserialize<List<SuggestionVisual>>(json);
                                        Console.WriteLine($"[Overlay] Parsed {suggestions?.Count ?? 0} suggestions");
                                        UpdateSuggestions(suggestions);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"[Overlay] JSON Error: {ex.Message}");
                                        System.Diagnostics.Debug.WriteLine($"Overlay JSON Error: {ex.Message}");
                                    }
                                });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Pipe Error: {ex.Message}");
                    await Task.Delay(1000);
                }
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Make window click-through initially or handle hit testing carefully
            // For this demo, we might want to keep it interactive for the popups
            // But the canvas area should be transparent to clicks if possible, 
            // except where we draw squiggles. 
            // Implementing full click-through with hole-punching is complex, 
            // so we'll stick to TopMost and assume the user clicks 'Scan' in Word to activate us 
            // or we poll.
            
            // In a real app, we'd use SetWindowLong to make it transparent to input 
            // when no popup is open, and toggle it when a squiggle is hovered/clicked.
        }

        // Public method to be called (via IPC/WCF/Pipe) to update suggestions
        public void UpdateSuggestions(List<SuggestionVisual> suggestions)
        {
            Console.WriteLine($"[Overlay] UpdateSuggestions called with {suggestions?.Count ?? 0} items");
            OverlayCanvas.Children.Clear();
            if (suggestions == null)
            {
                Console.WriteLine("[Overlay] Suggestions is null, returning");
                return;
            }
            
            int added = 0;
            foreach (var s in suggestions)
            {
                Console.WriteLine($"[Overlay] Adding squiggle at {s.Rect} for '{s.OriginalText}' -> '{s.suggestion}'");
                var path = SquiggleRenderer.CreateSquiggle(s.Rect, s.type);
                path.Tag = s;
                path.MouseLeftButtonDown += Squiggle_Click;
                path.MouseEnter += Squiggle_MouseEnter;
                path.MouseLeave += Squiggle_MouseLeave;
                OverlayCanvas.Children.Add(path);
                added++;
            }
            Console.WriteLine($"[Overlay] Added {added} squiggles to canvas");
            
            // Also update the static widget
            if (suggestions.Count > 0)
            {
                Console.WriteLine($"[Overlay] Updating widget with {suggestions.Count} suggestions");
                App.Widget?.UpdateSuggestions(suggestions);
            }
            else
            {
                Console.WriteLine("[Overlay] No suggestions, hiding widget");
                App.Widget?.Hide();
            }
        }

        private void Squiggle_MouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
        }
        
        private void Squiggle_MouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = null;
        }

        private void Squiggle_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Path path && path.Tag is SuggestionVisual s)
            {
                // Show popup with detailed information
                ErrorTypeText.Text = $"{char.ToUpper(s.type[0])}{s.type.Substring(1)} Error";
                OriginalText.Text = s.OriginalText;
                
                SuggestionPopup.PlacementTarget = path;
                SuggestionPopup.IsOpen = true;
                SuggestionsList.ItemsSource = new[] { s };
            }
        }

        private async void ApplySuggestion_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is SuggestionVisual s)
            {
                // Communicate back to Add-in to apply change
                SuggestionPopup.IsOpen = false;
                
                try
                {
                    // Send suggestion ID to Word add-in
                    await Task.Run(async () =>
                    {
                        using (var client = new System.IO.Pipes.NamedPipeClientStream(".", "WordOverlayAcceptPipe", System.IO.Pipes.PipeDirection.Out))
                        {
                            await client.ConnectAsync(2000);
                            using (var writer = new System.IO.StreamWriter(client))
                            {
                                await writer.WriteAsync(s.id);
                            }
                        }
                    });
                    
                    // Remove the squiggle from canvas
                    var pathToRemove = OverlayCanvas.Children.OfType<System.Windows.Shapes.Path>()
                        .FirstOrDefault(p => p.Tag is SuggestionVisual sv && sv.id == s.id);
                    if (pathToRemove != null)
                    {
                        OverlayCanvas.Children.Remove(pathToRemove);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error applying suggestion: {ex.Message}");
                    MessageBox.Show($"Failed to apply suggestion: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void Ignore_Click(object sender, RoutedEventArgs e)
        {
            SuggestionPopup.IsOpen = false;
        }
    }

    public class SuggestionVisual
    {
        public string type { get; set; }
        public string suggestion { get; set; }
        public Rect Rect { get; set; }
        public string OriginalText { get; set; }
        public string id { get; set; }
        public int from { get; set; }
        public int to { get; set; }
    }
}
