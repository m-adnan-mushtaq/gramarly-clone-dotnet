using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace WordOverlayProofreader.Overlay
{
    public partial class FloatingButton : Window
    {
        public event EventHandler ButtonClicked;
        
        private bool _isDragging = false;
        private Point _dragStart;
        private int _mistakeCount = 0;

        public FloatingButton()
        {
            InitializeComponent();
            PositionAtBottomRight();
            Console.WriteLine("[FloatingButton] Initialized");
        }

        private void PositionAtBottomRight()
        {
            var workArea = SystemParameters.WorkArea;
            this.Left = workArea.Right - this.Width - 20;
            this.Top = workArea.Bottom - this.Height - 20;
        }

        public void UpdateCount(int count)
        {
            _mistakeCount = count;
            CountText.Text = count.ToString();
            
            if (count > 0)
            {
                this.Show();
                Console.WriteLine($"[FloatingButton] Showing with count: {count}");
            }
            else
            {
                this.Hide();
                Console.WriteLine($"[FloatingButton] Hiding - no mistakes");
            }
        }

        private void Circle_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                _isDragging = true;
                _dragStart = e.GetPosition(this);
                ((Ellipse)sender).CaptureMouse();
                e.Handled = true;
            }
        }

        private void Circle_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDragging && e.LeftButton == MouseButtonState.Pressed)
            {
                var currentPos = e.GetPosition(this);
                var offset = currentPos - _dragStart;
                
                this.Left += offset.X;
                this.Top += offset.Y;
                
                e.Handled = true;
            }
        }

        private void Circle_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDragging)
            {
                _isDragging = false;
                ((Ellipse)sender).ReleaseMouseCapture();
                
                // If it was a click (not a drag), trigger the button event
                var currentPos = e.GetPosition(this);
                var distance = Math.Sqrt(Math.Pow(currentPos.X - _dragStart.X, 2) + Math.Pow(currentPos.Y - _dragStart.Y, 2));
                
                if (distance < 5) // It was a click
                {
                    Console.WriteLine($"[FloatingButton] Clicked! Opening sidebar...");
                    ButtonClicked?.Invoke(this, EventArgs.Empty);
                }
                
                e.Handled = true;
            }
        }
    }
}
