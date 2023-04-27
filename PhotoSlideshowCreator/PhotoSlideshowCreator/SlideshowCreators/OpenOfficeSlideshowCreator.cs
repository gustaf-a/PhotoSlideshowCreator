using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using PhotoSlideshowCreator.Data;
using PhotoSlideshowCreator.IO;

namespace PhotoSlideshowCreator.SlideshowCreators;

internal class OpenOfficeSlideshowCreator : ISlideshowCreator
{
    private static readonly string[] UnsupportedImageExtensions = new[] { ".webp" };

    public void CreateSlideshow(SourceData sourceData, SlideshowOptions slideshowOptions)
    {
        string outputPath = Path.Combine(slideshowOptions.OutputFolder, FileNameGenerator.GenerateUniqueFileName("slideshow", ".pptx"));

        using (PresentationDocument presentationDocument = PresentationDocument.Create(outputPath, PresentationDocumentType.Presentation))
        {
            // Create the presentation
            PresentationPart presentationPart = presentationDocument.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            CreateSlideMasterPart(presentationPart);

            SlideIdList slideIdList = presentationPart.Presentation.AppendChild(new SlideIdList());

            uint slideId = 1;

            foreach (var imageFile in sourceData.ImageFiles)
            {
                if (IsUnsupportedImageType(imageFile))
                    continue;

                // Create a new slide part and add it to the presentation
                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                string relationshipId = presentationPart.GetIdOfPart(slidePart);
                SlideId slideIdEntry = new SlideId() { Id = slideId, RelationshipId = relationshipId };
                slideIdList.Append(slideIdEntry);

                slidePart.Slide = CreateSlide(imageFile, slidePart, slideshowOptions);

                slideId++;

                presentationPart.Presentation.Save();
            }
        }
    }

    private static bool IsUnsupportedImageType(string imageFile)
    {
        foreach (var unsupportedExtension in UnsupportedImageExtensions)
            if(imageFile.Contains(unsupportedExtension))
                return true;

        return false;   
    }

    private static void CreateSlideMasterPart(PresentationPart presentationPart)
    {
        SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();

        var slideMaster = new SlideMaster();

        slideMaster.SlideLayoutIdList = new SlideLayoutIdList();

        slideMasterPart.SlideMaster = slideMaster;
        slideMasterPart.SlideMaster.Save();
    }

    private Slide CreateSlide(string imageFile, SlidePart slidePart, SlideshowOptions slideshowOptions)
    {
        var slide = new Slide(
            new CommonSlideData(
                new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() { Id = 1U, Name = "Slide" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new A.TransformGroup()),
                    CreateBackgroundShape(slideshowOptions.BackgroundColor),
                    CreatePictureShape(slidePart, imageFile))),
            new ColorMapOverride(new A.MasterColorMapping()),
            new Transition(new DocumentFormat.OpenXml.Presentation.FadeTransition())
            );

        slide.Timing = new Timing
        {
            TimeNodeList = new TimeNodeList(
                new CommonTimeNode
                {
                    Duration = $"PT{slideshowOptions.SlideDuration}S",
                    Restart = TimeNodeRestartValues.Never,
                    Fill = TimeNodeFillValues.Transition
                }
            )
        };

        return slide;
    }

    private Shape CreateBackgroundShape(System.Drawing.Color backgroundColor)
    {
        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties() { Id = 2U, Name = "Background" },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties()),
            new ShapeProperties(
                new A.Transform2D(new A.Offset(), new A.Extents()),
                new A.SolidFill(
                    new A.RgbColorModelHex()
                    {
                        Val = $"{backgroundColor.R:X2}{backgroundColor.G:X2}{backgroundColor.B:X2}"
                    })),
            new DocumentFormat.OpenXml.Presentation.ShapeStyle());
    }

    private Picture CreatePictureShape(SlidePart slidePart, string imageFile)
    {
        ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (FileStream stream = new FileStream(imageFile, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        string imageName = "Image";
        uint imageWidthEMU = 0;
        uint imageHeightEMU = 0;

        using (FileStream stream = new FileStream(imageFile, FileMode.Open))
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);
                memoryStream.Position = 0;

                using (System.Drawing.Image image = System.Drawing.Image.FromStream(memoryStream))
                {
                    imageWidthEMU = (uint)(image.Width * 9525);
                    imageHeightEMU = (uint)(image.Height * 9525);
                }
            }
        }


        var picture = new Picture(
            new NonVisualPictureProperties(
                new NonVisualDrawingProperties() { Id = UInt32Value.FromUInt32(4U), Name = imageName },
                new NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true, NoResize = true }),
                new ApplicationNonVisualDrawingProperties(new DocumentFormat.OpenXml.Presentation.PlaceholderShape())),
            new BlipFill(
                new A.Blip() { Embed = slidePart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                new A.Stretch(new A.FillRectangle())),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset() { X = 0L, Y = 0L },
                    new A.Extents() { Cx = imageWidthEMU, Cy = imageHeightEMU }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        return picture;
    }
}
