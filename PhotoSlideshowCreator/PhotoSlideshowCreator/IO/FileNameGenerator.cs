namespace PhotoSlideshowCreator.IO;

internal static class FileNameGenerator
{
    public static string GenerateUniqueFileName(string fileName, string fileExtension)
    {
        var dateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");

        return Path.GetFileNameWithoutExtension(fileName) + "_" + dateTime + fileExtension;
    }
}
