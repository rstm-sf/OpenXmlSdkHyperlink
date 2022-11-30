using System;
using System.IO;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using VerifyXunit;
using Xunit;

namespace OpenXmlSdkHyperlink;

[UsesVerify]
public class HyperlinkTests
{
    [Fact]
    public async Task Test()
    {
        var uri = new Uri("https://github.com/OfficeDev/Open-XML-SDK");

        using var docPackage = WordprocessingDocument.Create(new MemoryStream(), WordprocessingDocumentType.Document);
        var part = docPackage.AddMainDocumentPart();
        part.Document = new Document { Body = new Body() };

        var run = new Run(new Text("Hyperlink") { Space = SpaceProcessingModeValues.Preserve })
        {
            RunProperties = new RunProperties(
                new RunStyle { Val = "Hyperlink" },
                new Underline { Val = UnderlineValues.Single },
                new Color { ThemeColor = ThemeColorValues.Hyperlink })
        };
        var hyperlink = new Hyperlink(run)
        {
            Id = part.AddHyperlinkRelationship(uri, true).Id,
            History = true
        };
        part.Document.Body.AddChild(new Paragraph(hyperlink));

        docPackage.Save();

        var i = 1;
        foreach (var h in part.Document.Descendants<Hyperlink>())
        {
            h.Id = i.ToString("x8");
            i++;
        }

        await Verifier.Verify(FormatXml(part.Document.OuterXml));
    }

    private static readonly XmlWriterSettings XmlWriterSettings = new()
    {
        Indent = true,
        NewLineOnAttributes = false,
        IndentChars = "  ",
        ConformanceLevel = ConformanceLevel.Document
    };

    private static string FormatXml(string value)
    {
        using var textReader = new StringReader(value);
        using var xmlReader = XmlReader.Create(
            textReader, new XmlReaderSettings { ConformanceLevel = XmlWriterSettings.ConformanceLevel });
        using var textWriter = new StringWriter();
        using (var xmlWriter = XmlWriter.Create(textWriter, XmlWriterSettings))
            xmlWriter.WriteNode(xmlReader, true);
        return textWriter.ToString();
    }
}