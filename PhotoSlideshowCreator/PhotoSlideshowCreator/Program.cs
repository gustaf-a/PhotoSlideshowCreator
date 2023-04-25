using Microsoft.Office.Interop.PowerPoint;

namespace PhotoSlideshowCreator;

class Program
{
    private static readonly string[] ImageExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };

    private static readonly string[] YesInputs = new[] { "y", "yes" };

    static void Main(string[] args)
    {
        if (!GetSourceFolder(args, out string sourceFolder))
            return;

        var imageFiles = Directory.EnumerateFiles(sourceFolder)
            .Where(file => IsImage(file))
            .ToList();

        if (imageFiles.Count == 0)
        {
            Console.WriteLine("No images found in the current folder.");
            return;
        }

        Console.WriteLine($"Found {imageFiles.Count} image files.");

        Console.WriteLine($"Starting to create slideshow.");

        CreatePowerPoint(imageFiles, sourceFolder);

        Console.WriteLine("PowerPoint created successfully!");
    }

    private static bool GetSourceFolder(string[] args, out string sourceFolder)
    {
        sourceFolder = string.Empty;

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "-s" && i + 1 < args.Length)
            {
                sourceFolder = args[i + 1];
                break;
            }
        }

        if (IsExistingFolder(sourceFolder))
        {
            Console.WriteLine($"Using source folder '{sourceFolder}'.");
            return true;
        }

        Console.WriteLine("No valid source folder path provided.");
        Console.WriteLine("Paste in a folderpath");
        Console.WriteLine($"or press enter to use the current folder {Environment.CurrentDirectory}:");
        sourceFolder = Console.ReadLine();

        if (IsExistingFolder(sourceFolder))
        {
            Console.WriteLine($"Using source folder '{sourceFolder}'.");
            return true;
        }

        Console.WriteLine($"Do you want to use current folder '{Environment.CurrentDirectory}'? y/n");

        var rawYNInput = Console.ReadLine();
        if (YesInputs.Contains(rawYNInput.ToLower()))
        {
            sourceFolder = Environment.CurrentDirectory;

            Console.WriteLine($"Using source folder '{sourceFolder}'.");
            return true;
        }

        Console.WriteLine("Please restart and provide a valid sourcefolder using -s followed by path.");
        return false;
    }

    private static bool IsExistingFolder(string sourceFolder)
    {
        return !string.IsNullOrWhiteSpace(sourceFolder) && Directory.Exists(sourceFolder);
    }

    private static void CreatePowerPoint(IEnumerable<string> files, string sourceFolder)
    {
        var powerPointApp = new Application();

        var presentation = powerPointApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

        var slideLayout = presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];

        foreach (var file in files)
        {
            var slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, slideLayout);

            // Set the background color to black
            slide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
            slide.Background.Fill.BackColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            slide.Background.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            slide.Background.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.Background.Fill.Solid();

            if (IsImage(file))
            {
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
            }

            // Set slide transition
            slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            slide.SlideShowTransition.Duration = 1; // Transition duration in seconds
            slide.SlideShowTransition.AdvanceOnClick = Microsoft.Office.Core.MsoTriState.msoFalse;
            slide.SlideShowTransition.AdvanceOnTime = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.SlideShowTransition.AdvanceTime = 5; // Time before advancing to the next slide, in seconds
        }

        string outputPath = Path.Combine(sourceFolder, GenerateUniqueFileName("slideshow.pptx"));

        presentation.SaveAs(outputPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoTriStateMixed);

        presentation.Close();

        powerPointApp.Quit();
    }

    private static bool IsImage(string file)
    {
        return ImageExtensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
    }

    public static string GenerateUniqueFileName(string fileName)
    {
        var dateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");

        int fileNumber = 1;
        string newFileName = fileName;

        while (File.Exists(newFileName))
        {
            fileNumber++;
            newFileName = Path.GetFileNameWithoutExtension(fileName) + "_"  + dateTime + "_" + fileNumber + Path.GetExtension(fileName);
        }

        return newFileName;
    }
}
