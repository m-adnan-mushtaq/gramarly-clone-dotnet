using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace WordOverlayProofreader.Overlay
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public static SuggestionsSidebar Sidebar { get; private set; }
        public static FloatingButton FloatingBtn { get; private set; }
        
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // Allocate console for debugging
            AllocConsole();
            Console.WriteLine("========================================");
            Console.WriteLine("  Overlay Application Starting");
            Console.WriteLine("========================================");
            Console.WriteLine($"Started at: {DateTime.Now}");
            Console.WriteLine("This window should stay open...");
            Console.WriteLine();
            
            // Handle unhandled exceptions
            AppDomain.CurrentDomain.UnhandledException += (s, args) =>
            {
                Console.WriteLine($"[FATAL ERROR] {args.ExceptionObject}");
            };
            
            DispatcherUnhandledException += (s, args) =>
            {
                Console.WriteLine($"[UI ERROR] {args.Exception}");
                args.Handled = true;
            };
            
            // Create floating button (minimized state)
            Console.WriteLine("[App] Creating FloatingButton...");
            FloatingBtn = new FloatingButton();
            FloatingBtn.ButtonClicked += FloatingBtn_Clicked;
            FloatingBtn.Hide(); // Start hidden
            Console.WriteLine("[App] FloatingButton created");
            
            // Create the suggestions sidebar (expanded state)
            Console.WriteLine("[App] Creating SuggestionsSidebar...");
            Sidebar = new SuggestionsSidebar();
            Sidebar.SuggestionAccepted += Sidebar_SuggestionAccepted;
            Sidebar.SuggestionDismissed += Sidebar_SuggestionDismissed;
            Sidebar.Hide(); // Start hidden until we have suggestions
            Console.WriteLine("[App] SuggestionsSidebar created");
        }
        
        private void FloatingBtn_Clicked(object sender, EventArgs e)
        {
            Console.WriteLine("[App] Floating button clicked - opening sidebar");
            FloatingBtn.Hide();
            Sidebar.Show();
        }
        
        private void Sidebar_SuggestionAccepted(object sender, string suggestionId)
        {
            Console.WriteLine($"[App] ========================================");
            Console.WriteLine($"[App] ACCEPT BUTTON CLICKED!");
            Console.WriteLine($"[App] Suggestion ID: {suggestionId}");
            Console.WriteLine($"[App] ========================================");
            // Send acceptance back to Word via named pipe
            Task.Run(async () =>
            {
                try
                {
                    Console.WriteLine($"[App] Connecting to acceptance pipe...");
                    using (var client = new System.IO.Pipes.NamedPipeClientStream(".", "WordOverlayAcceptPipe", System.IO.Pipes.PipeDirection.Out))
                    {
                        await client.ConnectAsync(5000);
                        Console.WriteLine($"[App] Connected to acceptance pipe!");
                        using (var writer = new System.IO.StreamWriter(client) { AutoFlush = true })
                        {
                            await writer.WriteLineAsync(suggestionId);
                            Console.WriteLine($"[App] ✓ Sent acceptance to Word: {suggestionId}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[App] ERROR sending acceptance: {ex.Message}");
                    Console.WriteLine($"[App] Stack trace: {ex.StackTrace}");
                }
            });
        }
        
        private void Sidebar_SuggestionDismissed(object sender, string suggestionId)
        {
            Console.WriteLine($"[App] Suggestion dismissed: {suggestionId}");
            // Optionally send dismissal to Word or just ignore
        }
        
        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool AllocConsole();
    }
}
