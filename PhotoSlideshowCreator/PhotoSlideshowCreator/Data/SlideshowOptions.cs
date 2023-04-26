using System.Drawing;

namespace PhotoSlideshowCreator.Data;

internal class SlideshowOptions
{
    public string OutputFolder { get; set; }

    public int SlideDuration { get; set; } = 5;

    public Color BackgroundColor { get; set; } = Color.Black;
}
