using System;
using System.Threading;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordOverlayProofreader.Addin
{
    /// <summary>
    /// Console application to test the add-in functionality without VSTO
    /// </summary>
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("=== Word Overlay Proofreader Test ===");
            Console.WriteLine("Make sure:");
            Console.WriteLine("1. Overlay application is running");
            Console.WriteLine("2. SuggestionServer is running");
            Console.WriteLine("3. Word is open with a document");
            Console.WriteLine();

            try
            {
                // Try to connect to running Word instance
                Word.Application wordApp = null;
                try
                {
                    wordApp = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                    Console.WriteLine("✓ Connected to Word");
                }
                catch
                {
                    Console.WriteLine("✗ Could not connect to Word. Please open Word first.");
                    Console.WriteLine("Press any key to exit...");
                    Console.ReadKey();
                    return;
                }

                // Create add-in instance
                var addIn = new ThisAddIn();
                addIn.Application = wordApp;
                
                // Initialize
                addIn.ThisAddIn_Startup(null, EventArgs.Empty);
                Console.WriteLine("✓ Add-in initialized");

                Console.WriteLine();
                Console.WriteLine("Commands:");
                Console.WriteLine("  s - Scan document");
                Console.WriteLine("  a - Toggle auto-scan");
                Console.WriteLine("  q - Quit");
                Console.WriteLine();

                bool running = true;
                bool autoScan = false;

                while (running)
                {
                    if (Console.KeyAvailable)
                    {
                        var key = Console.ReadKey(true);
                        
                        switch (key.KeyChar)
                        {
                            case 's':
                            case 'S':
                                Console.WriteLine("Scanning document...");
                                addIn.ScanDocument();
                                break;
                                
                            case 'a':
                            case 'A':
                                autoScan = !autoScan;
                                addIn.SetAutoScan(autoScan);
                                Console.WriteLine($"Auto-scan: {(autoScan ? "ON" : "OFF")}");
                                break;
                                
                            case 'q':
                            case 'Q':
                                running = false;
                                Console.WriteLine("Shutting down...");
                                break;
                        }
                    }
                    
                    await Task.Delay(100);
                }

                addIn.ThisAddIn_Shutdown(null, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
