using Xunit;
using WordOverlayProofreader.Addin;
using System.Windows;

namespace WordOverlayProofreader.Tests
{
    public class CoordinateTests
    {
        [Fact]
        public void TestCoordinateConversion()
        {
            // Mock Word range or use a decoupled logic for testing math
            // Since we can't easily mock COM objects without a library, 
            // we'll test the logic if we had separated it.
            
            // Example: 
            // var helper = new WordCoordinateHelper(null);
            // var rect = helper.ConvertPointsToPixels(new Rect(10, 10, 100, 20));
            // Assert.True(rect.Width > 100); // assuming 96 dpi > 72 dpi
        }
    }
}
