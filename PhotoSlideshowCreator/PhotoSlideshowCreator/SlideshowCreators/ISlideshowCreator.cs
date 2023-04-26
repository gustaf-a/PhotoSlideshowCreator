using PhotoSlideshowCreator.Data;

namespace PhotoSlideshowCreator.SlideshowCreators;

internal interface ISlideshowCreator
{
    void CreateSlideshow(SourceData sourceData, SlideshowOptions slideshowOptions);
}
