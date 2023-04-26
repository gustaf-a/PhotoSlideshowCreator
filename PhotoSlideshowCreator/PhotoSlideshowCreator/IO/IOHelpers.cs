namespace PhotoSlideshowCreator.IO;

internal static class IOHelpers
{
    private static readonly string[] ImageExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };

    private static readonly string[] YesInputs = new[] { "y", "yes" };

    public static bool IsImage(string file)
    {
        return ImageExtensions.Any(ext => file.EndsWith(ext, StringComparison.OrdinalIgnoreCase));
    }

    public static bool IsExistingFolder(string sourceFolder)
    {
        return !string.IsNullOrWhiteSpace(sourceFolder) && Directory.Exists(sourceFolder);
    }

    public static bool GetYesInput()
    {
        var rawYNInput = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(rawYNInput))
            return false;

        return YesInputs.Contains(rawYNInput.ToLower());
    }

    public static IEnumerable<string> EnumerateFilesRecursively(string folderPath)
    {
        var files = Directory.EnumerateFiles(folderPath);

        foreach (var subfolder in Directory.EnumerateDirectories(folderPath))
        {
            files = files.Concat(EnumerateFilesRecursively(subfolder));
        }

        return files;
    }

    public static bool SelectEnumOption<TEnum>(Type enumType, out TEnum selectedOption, string prompt)
    {
        if (!enumType.IsEnum)
            throw new ArgumentException("Type must be an enum.");

        var enumValues = Enum.GetValues(enumType).Cast<TEnum>().ToList();

        Console.WriteLine("Please choose an option:");
        for (int i = 0; i < enumValues.Count; i++)
            Console.WriteLine($"{i + 1}. {enumValues[i]}");

        Console.WriteLine("0. Exit");

        int selectedIndex;
        while (true)
        {
            if (int.TryParse(Console.ReadLine(), out selectedIndex) && selectedIndex >= 0 && selectedIndex <= enumValues.Count)
                break;
            else
                Console.WriteLine("Invalid input. Please enter a valid number.");
        }

        if (selectedIndex == 0)
        {
            selectedOption = default(TEnum);
            return false;
        }
     
        selectedOption = enumValues[selectedIndex - 1];
        return true;
    }
}
