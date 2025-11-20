using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordOverlayProofreader.Addin
{
    public class WordCoordinateHelper
    {
        private Word.Application _app;

        public WordCoordinateHelper(Word.Application app)
        {
            _app = app;
        }

        public System.Windows.Rect GetScreenRect(Word.Range range)
        {
            try 
            {
                // Get the bounding rectangle for the range
                int left = 0, top = 0, width = 0, height = 0;
                
                // GetPoint returns coordinates relative to the window
                _app.ActiveWindow.GetPoint(out left, out top, out width, out height, range);
                
                // Convert to screen coordinates
                var window = _app.ActiveWindow;
                int windowLeft = window.Left;
                int windowTop = window.Top;
                
                // Calculate screen position
                int screenLeft = windowLeft + left;
                int screenTop = windowTop + top;
                
                Console.WriteLine($"[Coord] Window: ({windowLeft},{windowTop}), Relative: ({left},{top},{width},{height}), Screen: ({screenLeft},{screenTop})");
                
                // If dimensions are 0, use default height based on font size
                if (width == 0) width = 50;
                if (height == 0) height = 20;
                
                return new System.Windows.Rect(screenLeft, screenTop, width, height);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Coord] ERROR: {ex.Message}");
                // Return a default rect at top-left as fallback
                return new System.Windows.Rect(100, 100, 50, 20);
            }
        }
    }
}
