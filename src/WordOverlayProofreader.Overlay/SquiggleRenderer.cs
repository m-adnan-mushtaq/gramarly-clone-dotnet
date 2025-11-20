using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

namespace WordOverlayProofreader.Overlay
{
    public static class SquiggleRenderer
    {
        public static Path CreateSquiggle(Rect rect, string type)
        {
            Console.WriteLine($"[SquiggleRenderer] Creating underline at {rect} for type {type}");
            
            var path = new Path();
            path.Stroke = GetBrushForType(type);
            path.StrokeThickness = 2.5;
            path.Cursor = System.Windows.Input.Cursors.Hand;
            
            // Simple straight underline (like Grammarly)
            var geometry = new LineGeometry(
                new Point(rect.Left, rect.Bottom),
                new Point(rect.Right, rect.Bottom)
            );
            
            path.Data = geometry;
            path.ToolTip = $"{type} error - Click for suggestions";
            
            Console.WriteLine($"[SquiggleRenderer] Created underline from ({rect.Left},{rect.Bottom}) to ({rect.Right},{rect.Bottom})");
            return path;
        }

        private static Brush GetBrushForType(string type)
        {
            switch (type?.ToLower())
            {
                case "spelling": return new SolidColorBrush(Color.FromRgb(228, 0, 0)); // #E40000
                case "grammar": return new SolidColorBrush(Color.FromRgb(1, 151, 213)); // #0197D5
                case "acronym": return new SolidColorBrush(Color.FromRgb(246, 110, 18)); // #F66E12
                case "morphology": return new SolidColorBrush(Color.FromArgb(207, 50, 1, 213)); // #3201D5D0
                case "structure": return new SolidColorBrush(Color.FromRgb(243, 178, 0)); // #F3B200
                case "style": return new SolidColorBrush(Color.FromRgb(180, 73, 242)); // #B449F2
                case "formatting": return new SolidColorBrush(Color.FromRgb(7, 158, 82)); // #079E52
                case "dialect": return Brushes.Teal;
                case "quranic": return Brushes.Gold;
                case "synonyms": return Brushes.Olive;
                case "linguistic": return Brushes.Cyan;
                default: return new SolidColorBrush(Color.FromRgb(228, 0, 0)); // Default to red
            }
        }
    }
}
