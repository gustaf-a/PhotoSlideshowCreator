# Photo Slideshow Creator

## About this program

Photo Slideshow Creator is a C# program written by Gustaf Ahlner that generates a PowerPoint presentation containing images from a specified folder. 
Images are automatically resized and fitted based on their largest dimension, maintaining their aspect ratio. 
The program applies a fade transition to the slides, which advance automatically after a specified duration. 
The background color of each slide is set to black.

## How to use the standalone exe

1. Download the standalone `.exe` file provided in the release section or build it using the instructions in the "How to build the standalone exe" section below.
2. Copy the `.exe` file to the folder containing your imagesthat you want to include in the PowerPoint presentation.
3. Double-click the `.exe` file to run the program.
4. The program will create a PowerPoint presentation named `output.pptx` in the same folder, containing all imagesfrom the folder.

## How to build the standalone exe

1. Make sure you have the .NET 6.0 SDK installed on your computer. If you don't have it, download it from the official Microsoft website: https://dotnet.microsoft.com/download/dotnet
2. Open a command prompt or terminal window and navigate to your project folder.
3. Run the following command to create a standalone `.exe` file for the program:

dotnet publish -p:PublishSingleFile=true -r win-x64 -c Release --self-contained true


This command will build the project with the Release configuration for the Windows x64 platform and create a self-contained, single-file executable.


4. After the build process is complete, you can find the generated `.exe` file in the following folder within your project directory:

bin\Release\net6.0\win-x64\publish\


5. Copy the `.exe` file to the folder containing your images and videos, and follow the instructions in the "How to use the standalone exe" section above.
