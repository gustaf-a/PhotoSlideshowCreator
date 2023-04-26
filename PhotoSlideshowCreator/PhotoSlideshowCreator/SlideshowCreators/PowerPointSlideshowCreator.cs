using Microsoft.Office.Interop.PowerPoint;
using PhotoSlideshowCreator.Data;
using PhotoSlideshowCreator.IO;

namespace PhotoSlideshowCreator.SlideshowCreators;

internal class PowerPointSlideshowCreator : ISlideshowCreator
{
    public void CreateSlideshow(SourceData sourceData, SlideshowOptions slideshowOptions)
    {
        var powerPointApp = new Application();

        var presentation = powerPointApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

        var slideLayout = presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];

        foreach (var file in sourceData.ImageFiles)
        {
            var slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, slideLayout);

            // Set the background color to black
            slide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
            slide.Background.Fill.BackColor.RGB = System.Drawing.ColorTranslator.ToOle(slideshowOptions.BackgroundColor);
            slide.Background.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(slideshowOptions.BackgroundColor);
            slide.Background.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.Background.Fill.Solid();

            // Load the image to get its dimensions
            using (var image = System.Drawing.Image.FromFile(file))
            {
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;

                float imageWidth = image.Width;
                float imageHeight = image.Height;

                float scaleFactor = Math.Min(slideWidth / imageWidth, slideHeight / imageHeight);
                float newWidth = imageWidth * scaleFactor;
                float newHeight = imageHeight * scaleFactor;

                // Calculate position to center the image on the slide
                float positionX = (slideWidth - newWidth) / 2;
                float positionY = (slideHeight - newHeight) / 2;

                slide.Shapes.AddPicture(file, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, positionX, positionY, newWidth, newHeight);
            }

            // Set slide transition
            slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            slide.SlideShowTransition.Duration = 1; // Transition duration in seconds
            slide.SlideShowTransition.AdvanceOnClick = Microsoft.Office.Core.MsoTriState.msoFalse;
            slide.SlideShowTransition.AdvanceOnTime = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.SlideShowTransition.AdvanceTime = slideshowOptions.SlideDuration; // Time before advancing to the next slide, in seconds
        }

        string outputPath = Path.Combine(slideshowOptions.OutputFolder, FileNameGenerator.GenerateUniqueFileName("slideshow", "pptx"));

        presentation.SaveAs(outputPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoTriStateMixed);

        presentation.Close();

        powerPointApp.Quit();
    }
}
