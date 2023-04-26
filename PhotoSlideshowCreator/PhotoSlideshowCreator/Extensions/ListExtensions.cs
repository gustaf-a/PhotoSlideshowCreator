namespace PhotoSlideshowCreator.Extensions;

internal static class ListExtensions
{
    public static void Shuffle(this IList<string> list)
    {
        var random = new Random();

        for (int i = list.Count - 1; i > 0; i--)
        {
            int j = random.Next(i + 1);

            (list[j], list[i]) = (list[i], list[j]);
        }
    }
}
