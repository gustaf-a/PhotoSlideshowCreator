using PhotoSlideshowCreator.Data;
using PhotoSlideshowCreator.IO;

namespace PhotoSlideshowCreator.SlideshowCreators;

internal class OpenOfficeSlideshowCreator : ISlideshowCreator
{
    public void CreateSlideshow(SourceData sourceData, SlideshowOptions slideshowOptions)
    {
        //create presentation


        foreach (var imageFile in sourceData.ImageFiles)
        {
            //add slide

            //set backgroundcolor to slideshowOptions.BackgroundColor

            //calculate scaleFactor to fit image inside slide

            //calculate position to center image on slide

            //add picture in slide

            //set slide transition to fade 

            //set slide duration to slideshowOptions.SlideDuration
        }

        string outputPath = //slideshowOptions.OutputFolder combined with FileNameGenerator.GenerateUniqueFileName("slideshow.<replace-with-correct-ending>")


        //save presentation

        //close things
    }
}
