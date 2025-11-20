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
                if (range == null || _app.ActiveWindow == null)
                {
                    Console.WriteLine($"[Coord] ERROR: Range or window is null");
                    return System.Windows.Rect.Empty;
                }
                
                // Get the bounding rectangle for the range
                int left = 0, top = 0, width = 0, height = 0;
                
                // Collapse to start to get precise position
                var startRange = range.Duplicate;
                startRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                
                // Get position of start
                _app.ActiveWindow.GetPoint(out left, out top, out width, out height, startRange);
                
                // Get full range dimensions
                int rangeLeft = 0, rangeTop = 0, rangeWidth = 0, rangeHeight = 0;
                _app.ActiveWindow.GetPoint(out rangeLeft, out rangeTop, out rangeWidth, out rangeHeight, range);
                
                // Use the start position but the full width
                if (rangeWidth > 0) width = rangeWidth;
                if (rangeHeight > 0) height = rangeHeight;
                
                // Convert to screen coordinates
                var window = _app.ActiveWindow;
                int windowLeft = window.Left;
                int windowTop = window.Top;
                
                // Add window chrome offset (title bar, etc)
                int chromeOffsetX = 8; // Windows border
                int chromeOffsetY = window.Top + 100; // Title bar + ribbon approximate
                
                // Calculate screen position
                int screenLeft = windowLeft + left + chromeOffsetX;
                int screenTop = windowTop + top + chromeOffsetY;
                
                Console.WriteLine($"[Coord] Range '{range.Text?.Trim()}': Window({windowLeft},{windowTop}) + Relative({left},{top}) + Chrome({chromeOffsetX},{chromeOffsetY}) = Screen({screenLeft},{screenTop}) Size({width}x{height})");
                
                // Use reasonable defaults if dimensions are still 0
                if (width <= 0) width = Math.Max(range.Text?.Length ?? 5 * 8, 50); // Estimate based on text length
                if (height <= 0) height = 20;
                
                return new System.Windows.Rect(screenLeft, screenTop, width, height);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Coord] ERROR: {ex.Message}");
                Console.WriteLine($"[Coord] Stack: {ex.StackTrace}");
                return System.Windows.Rect.Empty;
            }
        }
    }
}
