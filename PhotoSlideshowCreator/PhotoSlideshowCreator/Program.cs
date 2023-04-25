using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.Threading.Channels;

namespace PhotoSlideshowCreator;

class Program
{
    private static readonly string[] ImageExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };
    private static readonly string[] VideoExtensions = new[] { ".mp4", ".avi", ".mkv", ".mov", ".wmv" };

    private static readonly string[] YesInputs = new[] { "y", "yes" };

    static void Main(string[] args)
    {
        if (!GetSourceFolder(args, out string sourceFolder))
            return;

        var files = Directory.EnumerateFiles(sourceFolder)
            .Where(file => IsImage(file) || IsVideo(file))
            .ToList();

        if (files.Count == 0)
        {
            Console.WriteLine("No images or videos found in the current folder.");
            return;
        }

        Console.WriteLine($"Found {files.Count} files. Starting to create slideshow.");

        CreatePowerPoint(files, sourceFolder);

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

        Console.WriteLine("Do you want to use this folder? y/n");

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

            if (IsImage(file))
            {
                slide.Shapes.AddPicture(file, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, slide.Master.Width, slide.Master.Height);
            }
            else if (IsVideo(file))
            {
                slide.Shapes.AddMediaObject2(file, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, slide.Master.Width, slide.Master.Height);
            }

            // Set slide transition
            slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            slide.SlideShowTransition.Duration = 1; // Transition duration in seconds
            slide.SlideShowTransition.AdvanceOnClick = Microsoft.Office.Core.MsoTriState.msoFalse;
            slide.SlideShowTransition.AdvanceOnTime = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.SlideShowTransition.AdvanceTime = 5; // Time before advancing to the next slide, in seconds
        }

        string outputPath = Path.Combine(sourceFolder, GenerateUniqueFileName("output.pptx"));

        presentation.SaveAs(outputPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoTriStateMixed);

        presentation.Close();

        powerPointApp.Quit();
    }

    private static bool IsImage(string file)
    {
        return ImageExtensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsVideo(string file)
    {
        return VideoExtensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
    }

    public static string GenerateUniqueFileName(string fileName)
    {
        int fileNumber = 1;
        string newFileName = fileName;

        while (File.Exists(newFileName))
        {
            fileNumber++;
            newFileName = Path.GetFileNameWithoutExtension(fileName) + "_" + fileNumber + Path.GetExtension(fileName);
        }

        return newFileName;
    }
}
