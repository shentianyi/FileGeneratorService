using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Svg;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using System.Drawing.Imaging; 

namespace Tek4.Highcharts.Exporting.MSDocumentGenerator
{
    public class PowerPointGenerator
    {
        private const int slideCx = 9144000;
        private const int slideCy = 6858000;
        // pptx xml relation ids
        private const string slideMasterRId = "rId1";
        private const string slidePartRId = "rId2";
        private const string themPartRId = "rId3";
        // private  const string imagePartRId = "rId3";
        /// <summary>
        /// not complete
        /// </summary>
        /// <param name="svgDocs"></param>
        /// <param name="stream"></param>
        public static void CreatePowerPointStream(List<SvgDocument> svgDocs, Stream stream)
        {
            throw new Exception("not complete");
        }

        /// <summary>
        /// create pptx
        /// using OpenSDK
        /// </summary>
        /// <param name="svgDocs"></param>
        /// <param name="stream"></param>
        public static void CreatePowerPointXStream(List<SvgDocument> svgDocs, Stream stream)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (PresentationDocument presentationDoc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation, true))
                {
                    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();
                    CreatePresentationTempalte(presentationPart);

                    CreateImageSlideParts(presentationPart, svgDocs);
                    DeleteSlide(presentationPart, 0);
                    presentationDoc.Close();
                    ms.Seek(0, SeekOrigin.Begin);
                    ms.WriteTo(stream);
                }
            }
        }

        private static void CreatePresentationTempalte(PresentationPart presentationPart)
        {
            // create template pptx
            SlideMasterIdList slideMasterIdList = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = slideMasterRId });
            SlideIdList slideIdList = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = slidePartRId });
            SlideSize slideSize = new SlideSize() { Cx = slideCx, Cy = slideCy, Type = SlideSizeValues.Screen4x3 };
            NotesSize notesSize = new NotesSize() { Cx = slideCy, Cy = slideCx };
            DefaultTextStyle defaultTextStyle = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList, slideIdList, slideSize, notesSize, defaultTextStyle);

            SlidePart templateSlidePart;
            SlideLayoutPart slideLayoutPart;
            SlideMasterPart slideMasterPart;
            ThemePart themePart;


            templateSlidePart = CreateSlidePart(presentationPart);
            slideLayoutPart = CreateSlideLayoutPart(templateSlidePart);
            slideMasterPart = CreateSlideMasterPart(slideLayoutPart);
            themePart = CreateTheme(slideMasterPart);

            slideMasterPart.AddPart(slideLayoutPart, slideMasterRId);
            presentationPart.AddPart(slideMasterPart, slideMasterRId);
            presentationPart.AddPart(themePart, themPartRId); 
        }

        private static SlidePart CreateSlidePart(PresentationPart presentationPart)
        {
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>(slidePartRId);
            slidePart.Slide = CreateSlide();
            return slidePart;
        }

        private static List<SlidePart> CreateImageSlideParts(PresentationPart presentationPart, List<SvgDocument> svgDocs)
        {
            int id = 256;
            string relId;
            SlideId newSlideId;
            SlideLayoutId newSlideLayoutId;
            uint uniqueId = GetMaxUniqueId(presentationPart);
            uint maxSlideId = GetMaxSlideId(presentationPart.Presentation.SlideIdList);
            // get first slide master part: template
            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
              
            List<SlidePart> slideParts = new List<SlidePart>();
            for (int i = 0; i < svgDocs.Count; i++)
            {
                id++;
                using (MemoryStream ms = new MemoryStream())
                {
                    using (System.Drawing.Bitmap image = svgDocs[i].Draw())
                    {
                        image.Save(ms, ImageFormat.Bmp);
                        ms.Seek(0, SeekOrigin.Begin);
                        relId = "rId" + id;
                        // add new slide part
                        SlidePart slidePart = presentationPart.AddNewPart<SlidePart>(relId);

                        // add image part to slide part
                        ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Bmp, relId);
                        imagePart.FeedData(ms);
                        // add image slide
                        CreateImageSlide(relId).Save(slidePart);

                        // add slide layout part to slide part
                        SlideLayoutPart slideLayoutPart = slidePart.AddNewPart<SlideLayoutPart>();
                        CreateSlideLayoutPart().Save(slideLayoutPart);
                        slideMasterPart.AddPart(slideLayoutPart);
                        slideLayoutPart.AddPart(slideMasterPart);
                        
                        uniqueId++;
                        newSlideLayoutId = new SlideLayoutId();
                        newSlideLayoutId.RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart);
                        newSlideLayoutId.Id = uniqueId;
                        slideMasterPart.SlideMaster.SlideLayoutIdList.Append(newSlideLayoutId);
                        
                        // add slide part to presentaion slide list
                        maxSlideId++;
                        newSlideId = new SlideId();
                        newSlideId.RelationshipId = relId;
                        newSlideId.Id = maxSlideId;
                        presentationPart.Presentation.SlideIdList.Append(newSlideId);
                    }
                }
            }
            slideMasterPart.SlideMaster.Save();
            return slideParts;
        }

        /// <summary>
        /// create slide
        /// </summary>
        /// <returns></returns>
        private static Slide CreateSlide()
        {
            return new Slide(
                        new CommonSlideData(
                            new ShapeTree(
                                new P.NonVisualGroupShapeProperties(
                                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                    new P.NonVisualGroupShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties()),
                                new GroupShapeProperties(new TransformGroup()),
                                new P.Shape(
                                    new P.NonVisualShapeProperties(
                                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                        new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                        new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                    new P.ShapeProperties(),
                                    new P.TextBody(
                                        new BodyProperties(),
                                        new ListStyle(),
                                        new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                        new ColorMapOverride(new MasterColorMapping()));
        }
        /// <summary>
        /// create image slide part
        /// </summary>
        /// <param name="imagePartRId"></param>
        /// <returns></returns>
        private static Slide CreateImageSlide(string imagePartRId)
        {
            return new Slide(
                   new CommonSlideData(
                     new ShapeTree(
                       new P.NonVisualGroupShapeProperties(
                         new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                         new P.NonVisualGroupShapeDrawingProperties(),
                         new ApplicationNonVisualDrawingProperties()),
                       new GroupShapeProperties(
                         new D.TransformGroup(
                           new D.Offset() { X = 0L, Y = 0L },
                           new D.Extents() { Cx = 0L, Cy = 0L },
                           new D.ChildOffset() { X = 0L, Y = 0L },
                           new D.ChildExtents() { Cx = 0L, Cy = 0L })),
                       new P.Picture(
                         new P.NonVisualPictureProperties(
                           new P.NonVisualDrawingProperties()
                           {
                               Id = (UInt32Value)4U,
                               Name = string.Empty,
                               Description = string.Empty
                           },
                           new P.NonVisualPictureDrawingProperties(
                             new D.PictureLocks() { NoChangeAspect = true }),
                           new ApplicationNonVisualDrawingProperties()),
                           new P.BlipFill(
                             new D.Blip() { Embed = imagePartRId },
                             new D.Stretch(
                               new D.FillRectangle())),
                           new P.ShapeProperties(
                             new D.Transform2D(
                               new D.Offset() { X = 0L, Y = 0L },
                               new D.Extents()
                               {
                                   Cx = slideCx,
                                   Cy = slideCy
                               }),
                             new D.PresetGeometry(
                               new D.AdjustValueList()) { Preset = D.ShapeTypeValues.Rectangle }
                           )))),
                   new ColorMapOverride(
                     new D.MasterColorMapping()));
        }

        /// <summary>
        /// create slide layout part by slide part
        /// </summary>
        /// <param name="slidePart"></param>
        /// <returns></returns>
        private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart.AddNewPart<SlideLayoutPart>(slideMasterRId);
            SlideLayout slideLayout = new SlideLayout(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph(new EndParagraphRunProperties()))))),
            new ColorMapOverride(new MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        /// <summary>
        /// create slide master part
        /// </summary>
        /// <param name="slideLayoutPart"></param>
        /// <returns></returns>
        private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart)
        {
            SlideMasterPart slideMasterPart = slideLayoutPart.AddNewPart<SlideMasterPart>(slideMasterRId);
            slideMasterPart.SlideMaster = new SlideMaster(
            new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()),
              new P.Shape(
              new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
              new P.ShapeProperties(),
              new P.TextBody(
                new BodyProperties(),
                new ListStyle(),
                new Paragraph())))),
            new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
            new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = slideMasterRId }),
            new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));

            return slideMasterPart;
        }
        /// <summary>
        /// create theme
        /// </summary>
        /// <param name="slideMasterPart"></param>
        /// <returns></returns>
        private static ThemePart CreateTheme(SlideMasterPart slideMasterPart)
        {
            ThemePart themePart = slideMasterPart.AddNewPart<ThemePart>(themPartRId);
            D.Theme theme = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
            new D.ColorScheme(
              new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
              new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
              new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
              new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
              new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
              new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
              new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
              new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
              new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
              new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
              new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
              new D.FontScheme(
              new D.MajorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" }),
              new D.MinorFont(
              new D.LatinFont() { Typeface = "Calibri" },
              new D.EastAsianFont() { Typeface = "" },
              new D.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
              new D.FormatScheme(
              new D.FillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                  new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                 new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 35000 },
                new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                 new D.SaturationModulation() { Val = 350000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 100000 }
                ),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.NoFill(),
              new D.PatternFill(),
              new D.GroupFill()),
              new D.LineStyleList(
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              },
              new D.Outline(
                new D.SolidFill(
                new D.SchemeColor(
                  new D.Shade() { Val = 95000 },
                  new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = D.LineCapValues.Flat,
                  CompoundLineType = D.CompoundLineValues.Single,
                  Alignment = D.PenAlignmentValues.Center
              }),
              new D.EffectStyleList(
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new D.EffectStyle(
                new D.EffectList(
                new D.OuterShadow(
                  new D.RgbColorModelHex(
                  new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new D.BackgroundFillStyleList(
              new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new D.GradientFill(
                new D.GradientStopList(
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
                new D.GradientStop(
                  new D.SchemeColor(new D.Tint() { Val = 50000 },
                    new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

            theme.Append(themeElements1);
            theme.Append(new D.ObjectDefaults());
            theme.Append(new D.ExtraColorSchemeList());

            themePart.Theme = theme;
            return themePart;
        }
        /// <summary>
        /// create slide layout part
        /// </summary>
        /// <returns></returns>
        private static SlideLayout CreateSlideLayoutPart()
        {
           return
                new SlideLayout(
                    new CommonSlideData(
                        new ShapeTree(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(
                                new D.TransformGroup(
                                    new D.Offset() { X = 0L, Y = 0L },
                                    new D.Extents() { Cx = 0L, Cy = 0L },
                                    new D.ChildOffset() { X = 0L, Y = 0L },
                                    new D.ChildExtents() { Cx = 0L, Cy = 0L })))
                    ) { Name = "Title Slide" },
                    new ColorMapOverride(
                        new D.MasterColorMapping())
                ) { Type = SlideLayoutValues.Title, Preserve = true }; 
        }
         
        /// <summary>
        /// get max slide id
        /// </summary>
        /// <param name="slideIdList"></param>
        /// <returns></returns>
        private static uint GetMaxSlideId(SlideIdList slideIdList)
        { 
            uint max = 256;
            if (slideIdList != null)
            {
                foreach (SlideId child in slideIdList.Elements<SlideId>())
                {
                    uint id = child.Id;
                    if (id > max)
                        max = id;
                }
            }
            return max;
        }

        /// <summary> 
        /// get max uniq id
        /// </summary>
        /// <param name="presentationPart"></param>
        /// <returns></returns>
        private static uint GetMaxUniqueId(PresentationPart presentationPart)
        {
            // Slide master identifiers have a minimum value of greater than or equal to 2147483648
            uint max = 2147483648;
            var slideMasterIdList = presentationPart.Presentation.SlideMasterIdList;
            if (slideMasterIdList != null)
            {
                // Get the maximum id value from the current set of children.
                foreach (SlideMasterId child in slideMasterIdList.Elements<SlideMasterId>())
                {
                    uint id = child.Id;
                    if (id > max)
                        max = id;
                }
            }

            foreach (var slideMasterPart in presentationPart.SlideMasterParts)
            {
                var slideLayoutIdList = slideMasterPart.SlideMaster.SlideLayoutIdList;
                if (slideLayoutIdList != null)
                {
                    // Get the maximum id value from the current set of children.
                    foreach (var child in slideLayoutIdList.Elements<SlideLayoutId>())
                    {
                        uint id = child.Id;
                        if (id > max)
                            max = id;
                    }
                }
            }
            return max;
        }

        /// <summary>
        /// Delete the specified slide from the presentation by slide index
        /// </summary>
        /// <param name="presentationPart"></param>
        /// <param name="slideIndex"></param>
        private static void DeleteSlide(PresentationPart presentationPart, int slideIndex)
        {  

            // Get the presentation from the presentation part.
            Presentation presentation = presentationPart.Presentation;

            // Get the list of slide IDs in the presentation.
            SlideIdList slideIdList = presentation.SlideIdList;

            // Get the slide ID of the specified slide
            SlideId slideId = slideIdList.ChildElements[slideIndex] as SlideId;

            // Get the relationship ID of the slide.
            string slideRelId = slideId.RelationshipId;

            // Remove the slide from the slide list.
            slideIdList.RemoveChild(slideId);

            //
            // Remove references to the slide from all custom shows.
            if (presentation.CustomShowList != null)
            {
                // Iterate through the list of custom shows.
                foreach (CustomShow customShow in presentation.CustomShowList.Elements<CustomShow>())
                {
                    if (customShow.SlideList != null)
                    {
                        // Declare a link list of slide list entries.
                        LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                        foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                        {
                            // Find the slide reference to remove from the custom show.
                            if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                            {
                                slideListEntries.AddLast(slideListEntry);
                            }
                        }

                        // Remove all references to the slide from the custom show.
                        foreach (SlideListEntry slideListEntry in slideListEntries)
                        {
                            customShow.SlideList.RemoveChild(slideListEntry);
                        }
                    }
                }
            }  
            // Get the slide part for the specified slide.
            SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

            // Remove the slide part.
            presentationPart.DeletePart(slidePart);
        }
    }
}
