# Microsoft Visio Stencil Creator

This application creates a Microsoft Visio stencil file (vssx) from PNG images without the need for Visio to be installed. It is a fork of [hoveytechllc/visio-stencil-creator](https://github.com/hoveytechllc/visio-stencil-creator), but updated with new capabilities such as a maximum image dimension to properly resize rectangular images and connector support on all images. 

I use this program heavily to make Visio stencils for architecture diagrams. See the section on [image sources](#image-sources) below.

Please, be sure to read the [additional notes](#additional-notes) section for any quirks or oddities about this program.

## Build Docker Image

The Docker image can be built with

```shell
docker buildx build -t kipidestan/visio-stencil-creator:<tag>
```

## Compile

You can compile the program for yourself with

```shell
dotnet restore /build/VisioStencilCreator.App/VisioStencilCreator.App.csproj
dotnet publish /build/VisioStencilCreator.App/VisioStencilCreator.App.csproj -o /app
```

## Usage

Docker image `kipidestan/visio-stencil-creator` prebuild using this source. Replace `<content-path>` with host path that contains images to be processed and where generated Visio stencil should be written to. 

```shell
docker run \
    -v <content-path>:/content \
    kipidestan/visio-stencil-creator:<tag> \
    "--image-path=/content" \
    "--image-pattern=*.png" \
    "--output-filename=/content/output.vssx"
```

## Additional Notes

* Parameter `--image-pattern` supports glob pattern searching, which internally uses [Microsoft.Extensions.FileSystemGlobbing](https://docs.microsoft.com/en-us/dotnet/api/microsoft.extensions.filesystemglobbing?view=aspnetcore-2.2) Nuget Package. e.g., --image-pattern=**/*.png will recursively search a directory for png files.
* While the code claims to support multiple image formats, I've only tested this fully with PNG.
* If you want to use SVGs, convert them to PNGs first with [rsvg-convert](https://github.com/GNOME/librsvg/blob/main/rsvg-convert.rst), [Inkscape](https://inkscape.org), or [ImageMagick's convert](https://imagemagick.org/script/convert.php).

## Image Sources

Below is a list of image sources I use when generating Stencils. Just make sure to attribute things correct.

* [Kubernetes Icons](https://github.com/kubernetes/community/tree/master/icons) - Kubernetes icons.
* [CNCF Artwork](https://github.com/cncf/artwork/tree/main?tab=readme-ov-file) - Cloud Native Computing Foundation logos.
* [Companies Logos](https://companieslogo.com/) - Company logos.
* [Amazing Icon Downloader](https://github.com/mattl-msft/Amazing-Icon-Downloader) - A browser plugin for Edge and Chrome that downloads SVGs from the Azure Portal.


