using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Validation;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
namespace Tek4.ConsoleTip
{
   public class PowerPointPicInstor
    {
       public void Insert(string filePath) {
           string newPresentation = filePath; 
           string imageFolder = @"D:\\";
           string[] imageFileExtensions =
             new[] { "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.png", "*.tif" }; 
           List<string> imageFileNames = GetImageFileNames(imageFolder,
             imageFileExtensions);
            
           if (imageFileNames.Count() > 0)
               CreateSlides(imageFileNames, newPresentation);
       }
        public void CreateSlides(List<string> imageFileNames,
            string newPresentation)
        {
            string relId;
            SlideId slideId;

            uint currentSlideId = 256;

            string imageFileNameNoPath;

            long imageWidthEMU = 0;
            long imageHeightEMU = 0;

            // Open the new presentation.
            using (PresentationDocument newDeck =
              PresentationDocument.Open(newPresentation, true))
            {
                PresentationPart presentationPart = newDeck.PresentationPart; 
                var slideMasterPart = presentationPart.SlideMasterParts.First(); 
                var slideLayoutPart = slideMasterPart.SlideLayoutParts.First(); 
                if (presentationPart.Presentation.SlideIdList == null)
                    presentationPart.Presentation.SlideIdList = new SlideIdList(); 
                foreach (string imageFileNameWithPath in imageFileNames)
                {
                    imageFileNameNoPath =
                      Path.GetFileNameWithoutExtension(imageFileNameWithPath); 
                    relId = "rel" + currentSlideId;
                     
                    ImagePartType imagePartType = ImagePartType.Png;
                    byte[] imageBytes = GetImageData(imageFileNameWithPath,ref imagePartType, ref imageWidthEMU, ref imageHeightEMU); 
                    var slidePart = presentationPart.AddNewPart<SlidePart>(relId);
                    GenerateSlidePart(imageFileNameNoPath, imageFileNameNoPath,imageWidthEMU, imageHeightEMU).Save(slidePart);

                    // Add the relationship between the slide and the
                    // slide layout.
                    slidePart.AddPart<SlideLayoutPart>(slideLayoutPart); 
                    var imagePart = slidePart.AddImagePart(ImagePartType.Jpeg,"relId12");
                    GenerateImagePart(imagePart, imageBytes);

                    // Add the new slide to the slide list.
                    slideId = new SlideId();
                    slideId.RelationshipId = relId;
                    slideId.Id = currentSlideId;
                    presentationPart.Presentation.SlideIdList.Append(slideId);

                    // Increment the slide id;
                    currentSlideId++;
                }

                // Save the changes to the slide master part.
                slideMasterPart.SlideMaster.Save();

                // Save the changes to the new deck.
                presentationPart.Presentation.Save();
            }
        }

        public  List<string> GetImageFileNames(string imageFolder,
          string[] imageFileExtensions)
        {
            // Create a list to hold the names of the files with the
            // requested extensions.
            List<string> fileNames = new List<string>();

            // Loop through each file extension.
            foreach (string extension in imageFileExtensions)
            {
                // Add all the files that match the current extension to the
                // list of file names.
                fileNames.AddRange(Directory.GetFiles(imageFolder, extension,
                  SearchOption.TopDirectoryOnly));
            }

            // Return the list of file names.
            return fileNames;
        }

        private   byte[] GetImageData(string imageFilePath,
          ref ImagePartType imagePartType, ref long imageWidthEMU,
          ref long imageHeightEMU)
        {
            byte[] imageFileBytes;
            using (FileStream fsImageFile = File.OpenRead(imageFilePath))
            {
                imageFileBytes = new byte[fsImageFile.Length];
                fsImageFile.Read(imageFileBytes, 0, imageFileBytes.Length);

                using (Bitmap imageFile = new Bitmap(fsImageFile))
                {
                    if (imageFile.RawFormat.Guid == ImageFormat.Bmp.Guid)
                        imagePartType = ImagePartType.Bmp;
                    else if (imageFile.RawFormat.Guid == ImageFormat.Gif.Guid)
                        imagePartType = ImagePartType.Gif;
                    else if (imageFile.RawFormat.Guid == ImageFormat.Jpeg.Guid)
                        imagePartType = ImagePartType.Jpeg;
                    else if (imageFile.RawFormat.Guid == ImageFormat.Png.Guid)
                        imagePartType = ImagePartType.Png;
                    else if (imageFile.RawFormat.Guid == ImageFormat.Tiff.Guid)
                        imagePartType = ImagePartType.Tiff;
                    else
                    {
                        throw new ArgumentException(
                          "Unsupported image file format: " + imageFilePath);
                    }

                    imageWidthEMU =
                    (long)
                    ((imageFile.Width / imageFile.HorizontalResolution) * 914400L);

                    imageHeightEMU =
                    (long)
                    ((imageFile.Height / imageFile.VerticalResolution) * 914400L);
                }
            }

            return imageFileBytes;
        }

        private   Slide GenerateSlidePart(string imageName,
          string imageDescription, long imageWidthEMU, long imageHeightEMU)
        {
            var element =
              new Slide(
                new CommonSlideData(
                  new ShapeTree(
                    new NonVisualGroupShapeProperties(
                      new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                      new NonVisualGroupShapeDrawingProperties(),
                      new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(
                      new D.TransformGroup(
                        new D.Offset() { X = 0L, Y = 0L },
                        new D.Extents() { Cx = 0L, Cy = 0L },
                        new D.ChildOffset() { X = 0L, Y = 0L },
                        new D.ChildExtents() { Cx = 0L, Cy = 0L })),
                    new Picture(
                      new NonVisualPictureProperties(
                        new NonVisualDrawingProperties()
                        {
                            Id = (UInt32Value)4U,
                            Name = imageName,
                            Description = imageDescription
                        },
                        new NonVisualPictureDrawingProperties(
                          new D.PictureLocks() { NoChangeAspect = true }),
                        new ApplicationNonVisualDrawingProperties()),
                        new BlipFill(
                          new D.Blip() { Embed = "relId12" },
                          new D.Stretch(
                            new D.FillRectangle())),
                        new ShapeProperties(
                          new D.Transform2D(
                            new D.Offset() { X = 0L, Y = 0L },
                            new D.Extents()
                            {
                                Cx = imageWidthEMU,
                                Cy = imageHeightEMU
                            }),
                          new D.PresetGeometry(
                            new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
                        )))),
                new ColorMapOverride(
                  new D.MasterColorMapping()));

            return element;
        }

        private   void GenerateImagePart(OpenXmlPart part,
          byte[] imageFileBytes)
        {
            // Write the contents of the image to the ImagePart.
            using (BinaryWriter writer = new BinaryWriter(part.GetStream()))
            {
                writer.Write(imageFileBytes);
                writer.Flush();
            }
        }
    }
}
