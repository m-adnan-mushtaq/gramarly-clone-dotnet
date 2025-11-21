using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WordOverlayProofreader.Overlay
{
    public partial class SuggestionWidget : Window
    {
        private List<SuggestionVisual> _currentSuggestions = new List<SuggestionVisual>();
        private int _currentIndex = 0;

        public event EventHandler<string> SuggestionAccepted;
        public event EventHandler<string> SuggestionDismissed;

        public SuggestionWidget()
        {
            InitializeComponent();
            PositionAtBottomRight();
            Loaded += (s, e) => PositionAtBottomRight();
        }

        private void PositionAtBottomRight()
        {
            var workingArea = SystemParameters.WorkArea;
            this.Left = workingArea.Right - this.Width - 20;
            this.Top = workingArea.Bottom - this.Height - 20;
        }

        public void UpdateSuggestions(List<SuggestionVisual> suggestions)
        {
            Console.WriteLine($"[Widget] UpdateSuggestions called with {suggestions?.Count ?? 0} items");
            _currentSuggestions = suggestions ?? new List<SuggestionVisual>();
            _currentIndex = 0;
            
            if (_currentSuggestions.Count > 0)
            {
                ShowSuggestion(_currentIndex);
                this.Show();
                this.Activate();
            }
            else
            {
                this.Hide();
            }
        }

        private void ShowSuggestion(int index)
        {
            if (index < 0 || index >= _currentSuggestions.Count)
            {
                this.Hide();
                return;
            }

            var suggestion = _currentSuggestions[index];
            
            // Update header with type-specific color
            ErrorTypeText.Text = FormatErrorType(suggestion.type);
            var typeColor = GetColorForType(suggestion.type);
            ErrorTypeText.Foreground = typeColor;
            ErrorTypeDot.Fill = typeColor;
            
            // Update content
            OriginalText.Text = suggestion.OriginalText;
            SuggestionText.Text = suggestion.suggestion;
            CountText.Text = $"{index + 1} of {_currentSuggestions.Count}";
            
            // Update navigation button states
            PreviousButton.IsEnabled = index > 0;
            NextButton.IsEnabled = index < _currentSuggestions.Count - 1;
            PreviousButton.Opacity = PreviousButton.IsEnabled ? 1.0 : 0.4;
            NextButton.Opacity = NextButton.IsEnabled ? 1.0 : 0.4;
        }

        private string FormatErrorType(string type)
        {
            if (string.IsNullOrEmpty(type)) return "Error";
            
            // Capitalize first letter
            return char.ToUpper(type[0]) + type.Substring(1).ToLower();
        }

        private Brush GetColorForType(string type)
        {
            switch (type?.ToLower())
            {
                case "spelling": 
                    return new SolidColorBrush(Color.FromRgb(228, 0, 0)); // #E40000 - Red
                case "grammar": 
                    return new SolidColorBrush(Color.FromRgb(21, 195, 154)); // #15C39A - Green
                case "style": 
                    return new SolidColorBrush(Color.FromRgb(180, 73, 242)); // #B449F2 - Purple
                case "acronym": 
                    return new SolidColorBrush(Color.FromRgb(246, 110, 18)); // #F66E12 - Orange
                case "morphology": 
                    return new SolidColorBrush(Color.FromRgb(50, 1, 213)); // #3201D5 - Blue
                case "structure": 
                    return new SolidColorBrush(Color.FromRgb(243, 178, 0)); // #F3B200 - Gold
                case "formatting": 
                    return new SolidColorBrush(Color.FromRgb(7, 158, 82)); // #079E52 - Dark Green
                default: 
                    return new SolidColorBrush(Color.FromRgb(1, 151, 213)); // #0197D5 - Blue
            }
        }

        private void AcceptButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine($"[Widget] ========================================");
            Console.WriteLine($"[Widget] AcceptButton_Click fired!");
            Console.WriteLine($"[Widget] Current index: {_currentIndex}");
            Console.WriteLine($"[Widget] Suggestions count: {_currentSuggestions?.Count ?? 0}");
            
            if (_currentSuggestions == null || _currentSuggestions.Count == 0)
            {
                Console.WriteLine($"[Widget] ERROR: No suggestions available");
                return;
            }
            
            if (_currentIndex < _currentSuggestions.Count)
            {
                var suggestion = _currentSuggestions[_currentIndex];
                Console.WriteLine($"[Widget] Accepting suggestion: '{suggestion.OriginalText}' -> '{suggestion.suggestion}'");
                Console.WriteLine($"[Widget] Suggestion ID: {suggestion.id}");
                Console.WriteLine($"[Widget] Invoking SuggestionAccepted event...");
                
                SuggestionAccepted?.Invoke(this, suggestion.id);
                
                Console.WriteLine($"[Widget] Event invoked, removing from list...");
                _currentSuggestions.RemoveAt(_currentIndex);
                
                if (_currentSuggestions.Count > 0)
                {
                    if (_currentIndex >= _currentSuggestions.Count)
                        _currentIndex = _currentSuggestions.Count - 1;
                    Console.WriteLine($"[Widget] Showing next suggestion...");
                    ShowSuggestion(_currentIndex);
                }
                else
                {
                    Console.WriteLine($"[Widget] No more suggestions, hiding widget");
                    this.Hide();
                }
            }
            Console.WriteLine($"[Widget] ========================================");
        }

        private async void DismissButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex < _currentSuggestions.Count)
            {
                var suggestion = _currentSuggestions[_currentIndex];
                Console.WriteLine($"[Widget] Dismissing suggestion: {suggestion.OriginalText}");
                
                // Send dismiss command to add-in
                await SendDismissToAddIn(suggestion.id);
                
                SuggestionDismissed?.Invoke(this, suggestion.id);
                
                _currentSuggestions.RemoveAt(_currentIndex);
                
                if (_currentSuggestions.Count > 0)
                {
                    if (_currentIndex >= _currentSuggestions.Count)
                        _currentIndex = _currentSuggestions.Count - 1;
                    ShowSuggestion(_currentIndex);
                }
                else
                {
                    this.Hide();
                }
            }
        }

        private async Task SendDismissToAddIn(string suggestionId)
        {
            try
            {
                await Task.Run(async () =>
                {
                    using (var client = new System.IO.Pipes.NamedPipeClientStream(".", "WordOverlayDismissPipe", System.IO.Pipes.PipeDirection.Out))
                    {
                        await client.ConnectAsync(2000);
                        using (var writer = new System.IO.StreamWriter(client))
                        {
                            await writer.WriteAsync(suggestionId);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Widget] Error sending dismiss: {ex.Message}");
            }
        }

        private void PreviousButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex > 0)
            {
                _currentIndex--;
                ShowSuggestion(_currentIndex);
            }
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex < _currentSuggestions.Count - 1)
            {
                _currentIndex++;
                ShowSuggestion(_currentIndex);
            }
        }
    }
}
