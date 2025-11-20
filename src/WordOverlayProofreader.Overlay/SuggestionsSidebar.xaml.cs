using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WordOverlayProofreader.Overlay
{
    public partial class SuggestionsSidebar : Window
    {
        public event EventHandler<string> SuggestionAccepted;
        public event EventHandler<string> SuggestionDismissed;
        
        private List<SuggestionVisual> _allSuggestions = new List<SuggestionVisual>();

        public SuggestionsSidebar()
        {
            InitializeComponent();
            PositionAtRightSide();
            Console.WriteLine("[Sidebar] Initialized");
        }

        private void PositionAtRightSide()
        {
            // Position at right edge of screen
            var workArea = SystemParameters.WorkArea;
            this.Left = workArea.Right - this.Width;
            this.Top = workArea.Top + 50;
            this.Height = workArea.Height - 100;
        }

        public void UpdateSuggestions(List<SuggestionVisual> suggestions)
        {
            Console.WriteLine($"[Sidebar] UpdateSuggestions called with {suggestions?.Count ?? 0} items");
            
            _allSuggestions = suggestions ?? new List<SuggestionVisual>();
            SuggestionsPanel.Children.Clear();

            if (_allSuggestions.Count == 0)
            {
                Console.WriteLine("[Sidebar] No suggestions, hiding");
                this.Hide();
                return;
            }

            Console.WriteLine($"[Sidebar] Building UI for {_allSuggestions.Count} suggestions");
            
            // Update header
            HeaderText.Text = $"Grammar Suggestions ({_allSuggestions.Count})";

            int index = 0;
            foreach (var suggestion in _allSuggestions)
            {
                var card = CreateSuggestionCard(suggestion, index);
                SuggestionsPanel.Children.Add(card);
                index++;
            }

            this.Show();
            Console.WriteLine("[Sidebar] Sidebar shown with all suggestions");
        }

        private Border CreateSuggestionCard(SuggestionVisual suggestion, int index)
        {
            var border = new Border
            {
                Background = Brushes.White,
                BorderBrush = GetBorderBrushForType(suggestion.type),
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(8),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(12)
            };

            var stackPanel = new StackPanel();

            // Header with type and number
            var headerPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            
            var numberText = new TextBlock
            {
                Text = $"#{index + 1}",
                FontWeight = FontWeights.Bold,
                FontSize = 12,
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 0, 8, 0)
            };
            
            var typeText = new TextBlock
            {
                Text = suggestion.type ?? "Error",
                FontWeight = FontWeights.Bold,
                FontSize = 13,
                Foreground = GetBrushForType(suggestion.type)
            };

            headerPanel.Children.Add(numberText);
            headerPanel.Children.Add(typeText);
            stackPanel.Children.Add(headerPanel);

            // Original text
            var originalLabel = new TextBlock
            {
                Text = "Original:",
                FontSize = 10,
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 0, 0, 2)
            };
            stackPanel.Children.Add(originalLabel);

            var originalBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(255, 235, 238)),
                Padding = new Thickness(8),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(0, 0, 0, 8)
            };
            
            var originalText = new TextBlock
            {
                Text = suggestion.OriginalText ?? "",
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(198, 40, 40)),
                TextWrapping = TextWrapping.Wrap
            };
            originalBorder.Child = originalText;
            stackPanel.Children.Add(originalBorder);

            // Suggestion text
            var suggestionLabel = new TextBlock
            {
                Text = "Suggestion:",
                FontSize = 10,
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 0, 0, 2)
            };
            stackPanel.Children.Add(suggestionLabel);

            var suggestionBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(232, 245, 233)),
                Padding = new Thickness(8),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(0, 0, 0, 10)
            };
            
            var suggestionText = new TextBlock
            {
                Text = suggestion.suggestion ?? "",
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(46, 125, 50)),
                TextWrapping = TextWrapping.Wrap
            };
            suggestionBorder.Child = suggestionText;
            stackPanel.Children.Add(suggestionBorder);

            // Buttons
            var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };

            var acceptButton = new Button
            {
                Content = "✓ Accept",
                Padding = new Thickness(15, 6, 15, 6),
                Margin = new Thickness(0, 0, 8, 0),
                Background = new SolidColorBrush(Color.FromRgb(76, 175, 80)),
                Foreground = Brushes.White,
                FontWeight = FontWeights.SemiBold,
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            acceptButton.Tag = suggestion.id;
            acceptButton.Click += AcceptButton_Click;

            var dismissButton = new Button
            {
                Content = "✗ Dismiss",
                Padding = new Thickness(15, 6, 15, 6),
                Background = new SolidColorBrush(Color.FromRgb(244, 67, 54)),
                Foreground = Brushes.White,
                FontWeight = FontWeights.SemiBold,
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            dismissButton.Tag = suggestion.id;
            dismissButton.Click += DismissButton_Click;

            buttonPanel.Children.Add(acceptButton);
            buttonPanel.Children.Add(dismissButton);
            stackPanel.Children.Add(buttonPanel);

            border.Child = stackPanel;
            return border;
        }

        private void AcceptButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var suggestionId = button?.Tag as string;
            
            Console.WriteLine($"[Sidebar] ========================================");
            Console.WriteLine($"[Sidebar] ACCEPT clicked for ID: {suggestionId}");
            
            if (!string.IsNullOrEmpty(suggestionId))
            {
                SuggestionAccepted?.Invoke(this, suggestionId);
                
                // Remove from list and rebuild UI
                _allSuggestions.RemoveAll(s => s.id == suggestionId);
                UpdateSuggestions(_allSuggestions);
            }
            Console.WriteLine($"[Sidebar] ========================================");
        }

        private void DismissButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var suggestionId = button?.Tag as string;
            
            Console.WriteLine($"[Sidebar] Dismiss clicked for ID: {suggestionId}");
            
            if (!string.IsNullOrEmpty(suggestionId))
            {
                SuggestionDismissed?.Invoke(this, suggestionId);
                
                // Remove from list and rebuild UI
                _allSuggestions.RemoveAll(s => s.id == suggestionId);
                UpdateSuggestions(_allSuggestions);
            }
        }

        private Brush GetBrushForType(string type)
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
                default: return Brushes.DarkGray;
            }
        }

        private Brush GetBorderBrushForType(string type)
        {
            var brush = GetBrushForType(type);
            if (brush is SolidColorBrush solidBrush)
            {
                var color = solidBrush.Color;
                var lighterColor = Color.FromArgb(100, color.R, color.G, color.B);
                return new SolidColorBrush(lighterColor);
            }
            return Brushes.LightGray;
        }
    }
}
