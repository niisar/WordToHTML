# WordToHTML

We will be using `OpenXML` and `OpenXmlPowerTools` to convert Word document into HTML.

### Step 1

Install Required Package

`Install-Package DocumentFormat.OpenXml`  

`Install-Package OpenXmlPowerTools`

### Add Reference

Right click in you Project in Solution Explorer  
then `Add >> Reference >> Select System.Drawing and WindowsBase` ![](https://4.bp.blogspot.com/-UXE7S7NUDCA/V5M1yuxnw7I/AAAAAAAAEpQ/WoOBV7g6r3Mox1YnXbaC40g1z6mLWbhcACLcB/s1600/2016-07-23%2B14_37_39-WordToHTML%2B-%2BMicrosoft%2BVisual%2BStudio.jpg)

### Follow the CODE Below
<pre>
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Drawing.Imaging;

namespace WordToHTML
{
    class Program
    {
        static void Main(string[] args)
        {
            byte[] byteArray = File.ReadAllBytes("kk.docx");

            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                {
                    int imageCounter = 0;
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = "My Page Title",
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo("img");
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }
                            if (imageFormat == null)
                                return null;

                            string imageFileName = "img/image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = HtmlConverter.ConvertToHtml(doc, settings);
                    File.WriteAllText("kk.html", html.ToStringNewLineOnAttributes());
                };
            }
        }
    }
}
</pre>