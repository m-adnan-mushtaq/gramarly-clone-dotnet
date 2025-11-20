using System;
using System.Collections.Generic;
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
            
            ErrorTypeText.Text = $"{char.ToUpper(suggestion.type[0])}{suggestion.type.Substring(1)} Error";
            ErrorTypeText.Foreground = GetColorForType(suggestion.type);
            OriginalText.Text = suggestion.OriginalText;
            SuggestionText.Text = suggestion.suggestion;
            CountText.Text = $"{index + 1} of {_currentSuggestions.Count}";
        }

        private Brush GetColorForType(string type)
        {
            switch (type.ToLower())
            {
                case "spelling": return Brushes.Red;
                case "grammar": return Brushes.Blue;
                case "style": return Brushes.Purple;
                default: return Brushes.Orange;
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

        private void DismissButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex < _currentSuggestions.Count)
            {
                var suggestion = _currentSuggestions[_currentIndex];
                Console.WriteLine($"[Widget] Dismissing suggestion: {suggestion.OriginalText}");
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
