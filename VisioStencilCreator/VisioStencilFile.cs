using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.IO.Packaging;
using System.Reflection;
using System.Text;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Formats.Bmp;
using SixLabors.ImageSharp.Metadata;

namespace VisioStencilCreator
{
    public class VisioStencilFile
    {
        public static Stream GenerateStencilFileFromImages(VisioStencilRequest request, VisioStencilConfig config)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            var templateStream = Assembly
                .GetExecutingAssembly()
                .GetManifestResourceStream("VisioStencilCreator.Resources.Template.vssx");

            if (templateStream == null)
                throw new Exception("Template stencil could not be loaded from resources.");

            var packageStream = new MemoryStream();
            templateStream.CopyTo(packageStream);
            packageStream.Seek(0, SeekOrigin.Begin);

            GenerateInternal(request.ImageFilePaths, packageStream, config);

            packageStream.Seek(0, SeekOrigin.Begin);

            return packageStream;
        }


        private static void GenerateInternal(IList<string> images, Stream packageStream, VisioStencilConfig config)
        {
            var package = Package.Open(
                packageStream,
                FileMode.Open,
                FileAccess.ReadWrite);

            var masterNames = string.Empty;
            var mastersXmlElements = string.Empty;

            var connections = string.Empty;
            foreach (VisioConnection connection in config.Connections)
            {
                var connectionRow = ConnectionRow
                    .Replace("{name}", connection.Name)
                    .Replace("{x}", connection.X)
                    .Replace("{y}", connection.Y);

                connections += ConnectionSection
                    .Replace("{connectionRows}", connectionRow);
            }

            var mastersPart = package.CreatePart(new Uri("/visio/masters/masters.xml", UriKind.Relative), "application/vnd.ms-visio.masters+xml");

            var sortedImages = images
                .OrderBy(path => Path.GetFileNameWithoutExtension(path), StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var image in sortedImages)
            {

                var id = images.IndexOf(image) + 1;
                // TODO: This assumes all files are a PNG. We should probably fix this in the future or only all pngs to be converted.
                var pngUri = new Uri($"/visio/media/image{id}.png", UriKind.Relative);
                var pngPart = package.CreatePart(pngUri, "image/png");

                double scaledWidth = 0.0,
                        scaledHeight = 0.0;

                using (Stream partStream = pngPart.GetStream(FileMode.Create,
                    FileAccess.ReadWrite))
                {
                    using (var fileStream = File.OpenRead(image))
                    {
                        // fileStream.CopyTo(partStream);

                        using var tempImage = Image.Load(fileStream);

                        double dpiX = 0.0, dpiY = 0.0;  // TODO: We need a failure if dpiX and Y are 0 after the if statement below
                        if (tempImage.Metadata.ResolutionUnits == PixelResolutionUnit.PixelsPerCentimeter)
                        {
                            dpiX = tempImage.Metadata.HorizontalResolution * 2.54f;
                            dpiY = tempImage.Metadata.VerticalResolution * 2.54;
                        }
                        else if (tempImage.Metadata.ResolutionUnits == PixelResolutionUnit.PixelsPerInch)
                        {
                            dpiX = tempImage.Metadata.HorizontalResolution;
                            dpiY = tempImage.Metadata.VerticalResolution;
                        }
                        else if (tempImage.Metadata.ResolutionUnits == PixelResolutionUnit.PixelsPerMeter)
                        {
                            dpiX = tempImage.Metadata.HorizontalResolution * 0.0254f;
                            dpiY = tempImage.Metadata.VerticalResolution * 0.0254f;
                        }
                        else
                        {
                            Console.WriteLine("Could not get proper image resolution metadata. Skipping file...");
                            continue;
                        }

                        int widthPx = tempImage.Width;
                        int heightPx = tempImage.Height;
                        int area = widthPx * heightPx;
                        double widthInches = widthPx / dpiX;
                        double heightInches = heightPx / dpiY;
                        double scale = Math.Sqrt((double)config.Image.TargetArea / area);
                        scaledWidth = widthInches * scale;
                        scaledHeight = heightInches * scale;

                        tempImage.SaveAsPng(partStream);
                    }
                }

                var masterUri = new Uri($"/visio/masters/master{id}.xml", UriKind.Relative);
                var masterPart = package.CreatePart(masterUri, "application/vnd.ms-visio.master+xml");

                var masterName = Path.GetFileNameWithoutExtension(image);

                using (Stream partStream = masterPart.GetStream(FileMode.Create,
                    FileAccess.ReadWrite))
                {
                    var masterXmlStencil = MasterXmlTemplate
                        .Replace("{fileName}", Path.GetFileName(image).ToString())
                        .Replace("{displayName}", masterName.ToString())
                        .Replace("{connections}", connections.ToString())
                        .Replace("{width}", scaledWidth.ToString())
                        .Replace("{height}", scaledHeight.ToString())
                        .Replace("{visAltText}", masterName.ToString());

                    using (var masterXmlStream = new MemoryStream(Encoding.UTF8.GetBytes(masterXmlStencil)))
                    {
                        masterXmlStream.CopyTo(partStream);
                    }
                }

                // TODO: The png extension assumes all file extensions are png. This shoudl be fixed in the future.
                masterPart.CreateRelationship(new Uri($"../media/image{id}.png", UriKind.Relative),
                    TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                    "rId1");

                masterNames += string.Format(MasterNameXmlTemplate, masterName);

                var imageThumbnail = ConvertImageToBase64Thumbnail(image);
                if (imageThumbnail == null)
                    continue;

                var masterXml = MastersMasterXmlTemplate
                    .Replace("{id}", id.ToString())
                    .Replace("{name}", masterName)
                    .Replace("{thumbnail}", imageThumbnail)
                    .Replace("{baseId}", $"{{{Guid.NewGuid().ToString()}}}")
                    .Replace("{uniqueId}", $"{{{Guid.NewGuid().ToString()}}}");

                mastersXmlElements += masterXml;

                mastersPart.CreateRelationship(masterUri,
                    TargetMode.Internal,
                    "http://schemas.microsoft.com/visio/2010/relationships/master",
                    $"rId{id}");

            }

            var propertiesPart = package.GetPart(new Uri("/docProps/app.xml", UriKind.Relative));

            using (Stream partStream = propertiesPart.GetStream(FileMode.Create,
                FileAccess.ReadWrite))
            {
                var propertiesXml = PropertiesXmlTemplate
                    .Replace("{masterCount}", images.Count.ToString())
                    .Replace("{partCount}", (images.Count + 1).ToString())
                    .Replace("{masterNames}", masterNames);

                using (var propertiesStream = new MemoryStream(Encoding.UTF8.GetBytes(propertiesXml)))
                {
                    propertiesStream.CopyTo(partStream);
                }
            }

            using (Stream partStream = mastersPart.GetStream(FileMode.Create,
                FileAccess.ReadWrite))
            {
                var propertiesXml = string.Format(MastersXmlTemplate, mastersXmlElements);

                using (var propertiesStream = new MemoryStream(Encoding.UTF8.GetBytes(propertiesXml)))
                {
                    propertiesStream.CopyTo(partStream);
                }
            }

            var documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));

            documentPart.CreateRelationship(new Uri("masters/masters.xml", UriKind.Relative),
                TargetMode.Internal,
                "http://schemas.microsoft.com/visio/2010/relationships/masters");

            package.Flush();
            package.Close();
        }

        private static string ConvertImageToBase64Thumbnail(string originalImagePath)
        {
            try
            {
                using var original = Image.Load<Rgba32>(originalImagePath);
                using var ms = new MemoryStream();
                original.Save(ms, new BmpEncoder());
                return Convert.ToBase64String(ms.ToArray());
            }
            catch (UnknownImageFormatException ex)
            {
                Console.WriteLine($"Unsupported image format: {ex.Message}.");
                return null;
            }
        }

        private const string MasterXmlTemplate = @"<?xml version='1.0' encoding='utf-8' ?>
<MasterContents xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xml:space='preserve'><Shapes>
<Shape ID='5' Type='Foreign' LineStyle='2' FillStyle='2' TextStyle='2'>
    <Cell N='PinX' V='3.49999985328088'/>
    <Cell N='PinY' V='6.49999974324154'/>
    <Cell N='Width' V='{width}'/>
    <Cell N='Height' V='{height}'/>
    <Cell N='LocPinX' F='Width*0.5'/>
    <Cell N='LocPinY' F='Height*0.5'/>
    <Cell N='Angle' V='0'/>
    <Cell N='FlipX' V='0'/>
    <Cell N='FlipY' V='0'/>
    <Cell N='ResizeMode' V='2'/>
    <Cell N='ImgOffsetX' V='0' F='ImgWidth*0'/>
    <Cell N='ImgOffsetY' V='0' F='ImgHeight*0'/>
    <Cell N='ImgWidth' F='Width*1'/>
    <Cell N='ImgHeight' F='Height*1'/>
    <Cell N='ClippingPath' V='' E='#N/A'/>
    <Cell N='EventDblClick' V='0' F='OPENTEXTWIN()'/>
    <Cell N='TxtPinX' V='0.3937007874015748' U='MM' F='Width*0.5'/>
    <Cell N='TxtPinY' V='-0.1476377952755905' U='MM' F='Height*-0.25'/>
    <Cell N='TxtWidth' V='1.764904432152344' F='TEXTWIDTH(TheText)'/>
    <Cell N='TxtHeight' V='0.1333828247070313' F='TEXTHEIGHT(TheText,TxtWidth)'/>
    <Cell N='TxtLocPinX' V='0' F='TxtWidth*0.5'/>
    <Cell N='TxtLocPinY' V='0' F='TxtHeight*0.5'/>
    <Cell N='TxtAngle' V='0'/>
    <Cell N='VerticalAlign' V='0'/>
    <Section N='User'>
        <Row N='visAltText'>
            <Cell N='Value' V='{visAltText}' U='STR'/>
        </Row>
    </Section>
    <Section N='Geometry' IX='0'><Cell N='NoFill' V='0'/><Cell N='NoLine' V='0'/><Cell N='NoShow' V='0'/><Cell N='NoSnap' V='0'/><Cell N='NoQuickDrag' V='0'/><Row T='RelMoveTo' IX='1'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row><Row T='RelLineTo' IX='2'><Cell N='X' V='1'/><Cell N='Y' V='0'/></Row><Row T='RelLineTo' IX='3'><Cell N='X' V='1'/><Cell N='Y' V='1'/></Row><Row T='RelLineTo' IX='4'><Cell N='X' V='0'/><Cell N='Y' V='1'/></Row><Row T='RelLineTo' IX='5'><Cell N='X' V='0'/><Cell N='Y' V='0'/></Row></Section><Section N='Property'><Row N='Label'><Cell N='Value' V='{fileName}' U='STR'/><Cell N='Prompt' V='{displayName}'/><Cell N='Label' V='{displayName}'/><Cell N='Format' V=''/><Cell N='SortKey' V=''/><Cell N='Type' V='0'/><Cell N='Invisible' V='0'/><Cell N='Verify' V='0'/><Cell N='DataLinked' V='0'/><Cell N='LangID' V='en-US'/><Cell N='Calendar' V='0'/></Row></Section>{connections}<ForeignData ForeignType='Bitmap' CompressionType='PNG'><Rel r:id='rId1'/></ForeignData></Shape></Shapes></MasterContents>";

        private const string ConnectionSection = @"<Section N='Connection'>{connectionRows}</Section>";
        private const string ConnectionRow = @"<Row T='Connection' N='{name}'>
    <Cell N='X' F='{x}'/>
    <Cell N='Y' F='{y}'/>
    <Cell N='DirX' V='0'/>
    <Cell N='DirY' V='0'/>
    <Cell N='Type' V='0'/>
    <Cell N='AutoGen' V='0'/>
    <Cell N='Prompt' V='' F='No Formula'/>
</Row>
    ";

        private const string MasterNameXmlTemplate = @"<vt:lpstr>{0}</vt:lpstr>";

        private const string PropertiesXmlTemplate = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Properties xmlns=""http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"" 
    xmlns:vt=""http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"">
    <Template></Template>
    <Application>Microsoft Visio</Application>
    <ScaleCrop>false</ScaleCrop>
    <HeadingPairs>
        <vt:vector size=""4"" baseType=""variant"">
            <vt:variant>
                <vt:lpstr>Pages</vt:lpstr>
            </vt:variant>
            <vt:variant>
                <vt:i4>1</vt:i4>
            </vt:variant>
            <vt:variant>
                <vt:lpstr>Masters</vt:lpstr>
            </vt:variant>
            <vt:variant>
                <vt:i4>{masterCount}</vt:i4>
            </vt:variant>
        </vt:vector>
    </HeadingPairs>
    <TitlesOfParts>
        <vt:vector size=""{partCount}"" baseType=""lpstr"">
            <vt:lpstr>Page-1</vt:lpstr>
            {masterNames}
        </vt:vector>
    </TitlesOfParts>
    <Manager></Manager>
    <Company></Company>
    <LinksUpToDate>false</LinksUpToDate>
    <SharedDoc>false</SharedDoc>
    <HyperlinkBase></HyperlinkBase>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>16.0000</AppVersion>
</Properties>";

        const string MastersXmlTemplate = @"<?xml version='1.0' encoding='utf-8' ?>
<Masters xmlns='http://schemas.microsoft.com/office/visio/2012/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xml:space='preserve'>{0}</Masters>";

        private const string MastersMasterXmlTemplate = @"<Master ID='{id}' NameU='{name}' IsCustomNameU='1' Name='{name}' IsCustomName='1' Prompt='' IconSize='1' AlignName='2' MatchByName='0' IconUpdate='1' UniqueID='{uniqueId}' BaseID='{baseId}' PatternFlags='0' Hidden='0' MasterType='2'><PageSheet LineStyle='0' FillStyle='0' TextStyle='0'><Cell N='PageWidth' V='8.5'/><Cell N='PageHeight' V='11'/><Cell N='ShdwOffsetX' V='0.125'/><Cell N='ShdwOffsetY' V='-0.125'/><Cell N='PageScale' V='1' U='IN_F'/><Cell N='DrawingScale' V='1' U='IN_F'/><Cell N='DrawingSizeType' V='0'/><Cell N='DrawingScaleType' V='0'/><Cell N='InhibitSnap' V='0'/><Cell N='PageLockReplace' V='0' U='BOOL'/><Cell N='PageLockDuplicate' V='0' U='BOOL'/><Cell N='UIVisibility' V='0'/><Cell N='ShdwType' V='0'/><Cell N='ShdwObliqueAngle' V='0'/><Cell N='ShdwScaleFactor' V='1'/><Cell N='DrawingResizeType' V='1'/></PageSheet><Icon>
{thumbnail}</Icon><Rel r:id='rId{id}'/></Master>";
    }
}
