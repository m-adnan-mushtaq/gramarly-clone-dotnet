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
            var path = new Path();
            path.Stroke = GetBrushForType(type);
            path.StrokeThickness = 2.0;
            path.Cursor = System.Windows.Input.Cursors.Hand;
            
            // Simple wave geometry - squiggly underline
            var geometry = new StreamGeometry();
            using (var ctx = geometry.Open())
            {
                ctx.BeginFigure(new Point(rect.Left, rect.Bottom), false, false);
                
                double waveLength = 4;
                double waveHeight = 2;
                
                for (double x = rect.Left; x < rect.Right; x += waveLength)
                {
                    ctx.BezierTo(
                        new Point(x + waveLength / 4, rect.Bottom - waveHeight),
                        new Point(x + waveLength * 3 / 4, rect.Bottom + waveHeight),
                        new Point(x + waveLength, rect.Bottom),
                        true, true);
                }
            }
            
            path.Data = geometry;
            path.ToolTip = $"{type} error - Click for suggestions";
            return path;
        }

        private static Brush GetBrushForType(string type)
        {
            switch (type.ToLower())
            {
                case "spelling": return Brushes.Red;
                case "grammar": return Brushes.Blue;
                case "style": return Brushes.Purple;
                case "dialect": return Brushes.Teal;
                case "morphology": return Brushes.Orange;
                case "structure": return Brushes.Brown;
                case "formatting": return Brushes.Gray;
                case "acronym": return Brushes.Green;
                case "quranic": return Brushes.Gold;
                case "synonyms": return Brushes.Olive;
                case "linguistic": return Brushes.Cyan;
                default: return Brushes.Red;
            }
        }
    }
}
