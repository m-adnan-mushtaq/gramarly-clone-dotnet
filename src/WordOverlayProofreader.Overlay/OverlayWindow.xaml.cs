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
            Console.WriteLine("[Overlay] Window_Loaded called");
            
            // Make the window cover the entire screen (all monitors)
            // This is crucial because coordinates from Word are in screen coordinates
            this.Left = SystemParameters.VirtualScreenLeft;
            this.Top = SystemParameters.VirtualScreenTop;
            this.Width = SystemParameters.VirtualScreenWidth;
            this.Height = SystemParameters.VirtualScreenHeight;
            
            Console.WriteLine($"[Overlay] Window positioned: ({this.Left}, {this.Top}) Size: {this.Width}x{this.Height}");
            
            // Make sure canvas matches window size
            OverlayCanvas.Width = this.Width;
            OverlayCanvas.Height = this.Height;
            
            Console.WriteLine($"[Overlay] Canvas sized: {OverlayCanvas.Width}x{OverlayCanvas.Height}");
            
            SetupVisibilityTimer();
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
            
            // Get current DPI scale
            var source = PresentationSource.FromVisual(this);
            double dpiX = 1.0;
            double dpiY = 1.0;
            
            if (source != null)
            {
                dpiX = source.CompositionTarget.TransformToDevice.M11;
                dpiY = source.CompositionTarget.TransformToDevice.M22;
                Console.WriteLine($"[Overlay] DPI Scale: {dpiX}x{dpiY}");
            }
            
            int added = 0;
            foreach (var s in suggestions)
            {
                // Convert screen pixels (from Word) to WPF DIPs
                var pixelRect = s.Rect;
                var dipRect = new Rect(
                    pixelRect.X / dpiX,
                    pixelRect.Y / dpiY,
                    pixelRect.Width / dpiX,
                    pixelRect.Height / dpiY
                );
                
                Console.WriteLine($"[Overlay] Adding squiggle at {dipRect} (Pixels: {pixelRect}) for '{s.OriginalText}'");
                var path = SquiggleRenderer.CreateSquiggle(dipRect, s.type);
                path.Tag = s;
                path.MouseLeftButtonDown += Squiggle_Click;
                path.MouseEnter += Squiggle_MouseEnter;
                path.MouseLeave += Squiggle_MouseLeave;
                OverlayCanvas.Children.Add(path);
                added++;
            }
            Console.WriteLine($"[Overlay] Added {added} squiggles to canvas");
            
            // Update both floating button and sidebar with all suggestions
            if (suggestions.Count > 0)
            {
                Console.WriteLine($"[Overlay] Showing floating button with {suggestions.Count} suggestions");
                App.FloatingBtn?.UpdateCount(suggestions.Count);
                
                // Pre-load sidebar with data but keep it hidden
                App.Sidebar?.UpdateSuggestions(suggestions);
                App.Sidebar?.Hide();
            }
            else
            {
                Console.WriteLine("[Overlay] No suggestions, hiding all widgets");
                App.FloatingBtn?.Hide();
                App.Sidebar?.Hide();
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
                // OriginalText.Text = s.OriginalText; // Removed in new UI
                
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

        private void CloseTooltip_Click(object sender, RoutedEventArgs e)
        {
            SuggestionPopup.IsOpen = false;
        }

        private void Ignore_Click(object sender, RoutedEventArgs e)
        {
            SuggestionPopup.IsOpen = false;
        }

        // P/Invoke for window visibility
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private System.Windows.Threading.DispatcherTimer _visibilityTimer;

        private void SetupVisibilityTimer()
        {
            _visibilityTimer = new System.Windows.Threading.DispatcherTimer();
            _visibilityTimer.Interval = TimeSpan.FromMilliseconds(500);
            _visibilityTimer.Tick += VisibilityTimer_Tick;
            _visibilityTimer.Start();
        }

        private void VisibilityTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                IntPtr hWnd = GetForegroundWindow();
                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);

                var process = System.Diagnostics.Process.GetProcessById((int)processId);
                if (process.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
                {
                    if (this.Visibility != Visibility.Visible)
                    {
                        this.Visibility = Visibility.Visible;
                        Console.WriteLine("[Overlay] Word is foreground, showing overlay");
                    }
                }
                else
                {
                    // Check if the foreground window is OUR overlay or sidebar or popup
                    // If we are interacting with the overlay, we shouldn't hide it
                    bool isOurApp = process.Id == System.Diagnostics.Process.GetCurrentProcess().Id;
                    
                    if (!isOurApp && this.Visibility == Visibility.Visible)
                    {
                        this.Visibility = Visibility.Hidden;
                        Console.WriteLine($"[Overlay] Word not foreground (Active: {process.ProcessName}), hiding overlay");
                    }
                }
            }
            catch (Exception)
            {
                // Ignore errors (process might have exited)
            }
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
