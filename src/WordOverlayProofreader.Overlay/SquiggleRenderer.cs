using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

namespace WordOverlayProofreader.Overlay
{
    public static class SquiggleRenderer
    {
        public static UIElement CreateSquiggle(Rect rect, string type)
        {
            Console.WriteLine($"[SquiggleRenderer] Creating underline at {rect} for type {type}");
            
            // Create a container to hold both the visual line and hit area
            var container = new System.Windows.Controls.Canvas();
            container.Width = rect.Width;
            container.Height = 12; // Increased height for better click area
            
            // Create the visible underline (bold like Grammarly)
            var path = new Path();
            path.Stroke = GetBrushForType(type);
            path.StrokeThickness = 3.5; // Bolder line like Grammarly
            path.StrokeDashArray = null; // Solid line
            path.IsHitTestVisible = false; // Let the container handle hits
            
            // Create straight underline
            var pathFigure = new PathFigure();
            pathFigure.StartPoint = new Point(0, 0);
            pathFigure.Segments.Add(new LineSegment(new Point(rect.Width, 0), true));
            
            var pathGeometry = new PathGeometry();
            pathGeometry.Figures.Add(pathFigure);
            path.Data = pathGeometry;
            
            // Add path to container
            container.Children.Add(path);
            
            // Create an invisible hit area above the line for easier clicking
            var hitArea = new System.Windows.Shapes.Rectangle();
            hitArea.Width = rect.Width;
            hitArea.Height = 12; // Large click area
            hitArea.Fill = Brushes.Transparent;
            hitArea.Cursor = System.Windows.Input.Cursors.Hand;
            hitArea.IsHitTestVisible = true;
            
            System.Windows.Controls.Canvas.SetTop(hitArea, -6); // Center the hit area around the line
            container.Children.Add(hitArea);
            
            // Position container on canvas
            System.Windows.Controls.Canvas.SetLeft(container, rect.Left);
            System.Windows.Controls.Canvas.SetTop(container, rect.Bottom - 1); // Slight adjustment for visual alignment
            System.Windows.Controls.Canvas.SetZIndex(container, 1000);
            
            container.Cursor = System.Windows.Input.Cursors.Hand;
            container.IsHitTestVisible = true;
            
            Console.WriteLine($"[SquiggleRenderer] Created bold underline with hit area at ({rect.Left},{rect.Bottom}) width {rect.Width}");
            return container;
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
