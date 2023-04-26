using PhotoSlideshowCreator.Data;
using PhotoSlideshowCreator.IO;
using PhotoSlideshowCreator.SlideshowCreators;
using PhotoSlideshowCreator.Extensions;

namespace PhotoSlideshowCreator;

class Program
{
    public enum SlideShowTypes
    {
        PowerPoint,
        OpenOffice
    };

    static void Main(string[] args)
    {
        SourceData sourceData = new ();
        SlideshowOptions slideshowOptions = new();

        GetSourceFiles(sourceData, args);

        PreProcessing(sourceData, slideshowOptions);

        CreateSlideShow(sourceData, slideshowOptions);
    }

    private static void GetSourceFiles(SourceData sourceData, string[] args)
    {
        if (!GetSourceFolder(args, out string sourceFolder))
            return;

        sourceData.SourceFolder = sourceFolder;

        var imageFiles = IOHelpers.EnumerateFilesRecursively(sourceFolder)
            .Where(file => IOHelpers.IsImage(file))
            .ToList();

        if (imageFiles.Count == 0)
        {
            Console.WriteLine("No images found in the current folder.");
            return;
        }

        sourceData.ImageFiles = imageFiles;

        Console.WriteLine($"Found {imageFiles.Count} image files.");
    }

    private static void PreProcessing(SourceData sourceData, SlideshowOptions slideshowOptions)
    {
        Console.WriteLine($"Do you want to shuffle the found files? y/n");
        if (IOHelpers.GetYesInput())
            sourceData.ImageFiles.Shuffle();
    }

    private static void CreateSlideShow(SourceData sourceData, SlideshowOptions slideshowOptions)
    {
        if (!IOHelpers.SelectEnumOption(typeof(SlideShowTypes), out SlideShowTypes selectedOption, "Select which slideshow you want to create:"))
            return;

        ISlideshowCreator creator = selectedOption switch
        {
            SlideShowTypes.PowerPoint => new PowerPointSlideshowCreator(),
            SlideShowTypes.OpenOffice => new OpenOfficeSlideshowCreator(),
            _ => null,
        };

        if (creator == null)
        {
            Console.WriteLine("No slideshow creator selected. Exiting.");
            return;
        }

        Console.WriteLine($"Starting to create slideshow.");

        creator.CreateSlideshow(sourceData, slideshowOptions);

        Console.WriteLine("Slideshow created successfully!");
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

        if (IOHelpers.IsExistingFolder(sourceFolder))
        {
            Console.WriteLine($"Using source folder '{sourceFolder}'.");
            return true;
        }

        Console.WriteLine("No valid source folder path provided.");
        Console.WriteLine("Paste in a folderpath");
        Console.WriteLine($"or press enter to use the current folder {Environment.CurrentDirectory}:");

        sourceFolder = Console.ReadLine();

        if (IOHelpers.IsExistingFolder(sourceFolder))
        {
            Console.WriteLine($"Using source folder '{sourceFolder}'.");
            return true;
        }

        Console.WriteLine($"Do you want to use current folder '{Environment.CurrentDirectory}'? y/n");
        if (IOHelpers.GetYesInput())
        {
            sourceFolder = Environment.CurrentDirectory;

            Console.WriteLine($"Using source folder '{sourceFolder}'.");
            return true;
        }

        Console.WriteLine("Please restart and provide a valid sourcefolder using -s followed by path.");
        return false;
    }
}
