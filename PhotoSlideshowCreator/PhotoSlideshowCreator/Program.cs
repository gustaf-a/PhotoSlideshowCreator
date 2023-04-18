using Microsoft.Office.Interop.PowerPoint;

namespace PhotoSlideshowCreator;

class Program
{
    private static readonly string[] ImageExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };
    private static readonly string[] VideoExtensions = new[] { ".mp4", ".avi", ".mkv", ".mov", ".wmv" };

    static void Main(string[] args)
    {
        string currentFolder1 = Directory.GetCurrentDirectory();
        string currentFolder = @"C:\PrivateRepos\PhotoSlideshowCreator\TestImages";

        var files = Directory.EnumerateFiles(currentFolder)
            .Where(file => IsImage(file) || IsVideo(file))
            .ToList();

        if (files.Count == 0)
        {
            Console.WriteLine("No images or videos found in the current folder.");
            return;
        }

        Console.WriteLine($"Found {files.Count} files. Starting to create slideshow.");

        CreatePowerPoint(files, currentFolder);

        Console.WriteLine("PowerPoint created successfully!");
    }

    private static void CreatePowerPoint(IEnumerable<string> files, string currentFolder)
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
        }

        string outputPath = Path.Combine(currentFolder, GenerateUniqueFileName("output.pptx"));
        
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
