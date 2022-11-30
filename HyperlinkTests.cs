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
        var documentPart = docPackage.AddMainDocumentPart();
        documentPart.Document = new Document {Body = new Body()};

        var numberPart = documentPart.AddNewPart<NumberingDefinitionsPart>();
        numberPart.Numbering = new Numbering();
        numberPart.Numbering.Save(documentPart.NumberingDefinitionsPart!);

        const int abstractId = 1;
        const int numberId = 1;
        var abstractNum = GetBulleted();
        abstractNum.AbstractNumberId = abstractId;
        var abstractNumId = new AbstractNumId { Val = abstractId };
        var numberingInstance = new NumberingInstance(abstractNumId) { NumberID = numberId };
        numberPart.Numbering.Append(numberingInstance, abstractNum);

        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Append(new ParagraphStyleId { Val = "ListParagraph" });
        paragraphProperties.Append(
            new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = numberId }));
        var itemParagraph = new Paragraph();
        itemParagraph.Append(paragraphProperties);
        itemParagraph.Append(
            new Run(
                new RunProperties(),
                new Text("Item") { Space = SpaceProcessingModeValues.Preserve }));

        var run = new Run(new Text("Hyperlink") {Space = SpaceProcessingModeValues.Preserve})
        {
            RunProperties = new RunProperties(
                new RunStyle {Val = "Hyperlink"},
                new Underline {Val = UnderlineValues.Single},
                new Color {ThemeColor = ThemeColorValues.Hyperlink})
        };
        var hyperlink = new Hyperlink(run)
        {
            Id = documentPart.AddHyperlinkRelationship(uri, true).Id,
            History = true
        };
        itemParagraph.Append(hyperlink);

        documentPart.Document.Body.AppendChild(itemParagraph);
        docPackage.Save();

        var i = 1;
        foreach (var h in documentPart.Document.Descendants<Hyperlink>())
        {
            h.Id = i.ToString("x8");
            i++;
        }

        await Verifier.Verify(FormatXml(documentPart.Document.OuterXml));
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
            textReader, new XmlReaderSettings {ConformanceLevel = XmlWriterSettings.ConformanceLevel});
        using var textWriter = new StringWriter();
        using (var xmlWriter = XmlWriter.Create(textWriter, XmlWriterSettings))
            xmlWriter.WriteNode(xmlReader, true);
        return textWriter.ToString();
    }

    private static AbstractNum GetBulleted()
    {
        var abstractNum1 = new AbstractNum {AbstractNumberId = 7};
        abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak",
            "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
        var nsid1 = new Nsid {Val = "433F0C6C"};
        var multiLevelType1 = new MultiLevelType {Val = MultiLevelValues.HybridMultilevel};
        var templateCode1 = new TemplateCode {Val = "934E79A6"};

        var level1 = new Level {LevelIndex = 0, TemplateCode = "04090001"};
        var startNumberingValue1 = new StartNumberingValue {Val = 1};
        var numberingFormat1 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText1 = new LevelText {Val = "·"};
        var levelJustification1 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties1 = new PreviousParagraphProperties();
        var indentation1 = new Indentation {Left = "720", Hanging = "360"};

        previousParagraphProperties1.Append(indentation1);

        var numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
        var runFonts1 = new RunFonts {Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol"};

        numberingSymbolRunProperties1.Append(runFonts1);

        level1.Append(startNumberingValue1);
        level1.Append(numberingFormat1);
        level1.Append(levelText1);
        level1.Append(levelJustification1);
        level1.Append(previousParagraphProperties1);
        level1.Append(numberingSymbolRunProperties1);

        var level2 = new Level {LevelIndex = 1, TemplateCode = "04090003"};
        var startNumberingValue2 = new StartNumberingValue {Val = 1};
        var numberingFormat2 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText2 = new LevelText {Val = "o"};
        var levelJustification2 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties2 = new PreviousParagraphProperties();
        var indentation2 = new Indentation {Left = "1440", Hanging = "360"};

        previousParagraphProperties2.Append(indentation2);

        var numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
        var runFonts2 = new RunFonts
        {
            Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New",
            ComplexScript = "Courier New"
        };

        numberingSymbolRunProperties2.Append(runFonts2);

        level2.Append(startNumberingValue2);
        level2.Append(numberingFormat2);
        level2.Append(levelText2);
        level2.Append(levelJustification2);
        level2.Append(previousParagraphProperties2);
        level2.Append(numberingSymbolRunProperties2);

        var level3 = new Level {LevelIndex = 2, TemplateCode = "04090005", Tentative = true};
        var startNumberingValue3 = new StartNumberingValue {Val = 1};
        var numberingFormat3 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText3 = new LevelText {Val = "§"};
        var levelJustification3 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties3 = new PreviousParagraphProperties();
        var indentation3 = new Indentation {Left = "2160", Hanging = "360"};

        previousParagraphProperties3.Append(indentation3);

        var numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
        var runFonts3 = new RunFonts {Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings"};

        numberingSymbolRunProperties3.Append(runFonts3);

        level3.Append(startNumberingValue3);
        level3.Append(numberingFormat3);
        level3.Append(levelText3);
        level3.Append(levelJustification3);
        level3.Append(previousParagraphProperties3);
        level3.Append(numberingSymbolRunProperties3);

        var level4 = new Level {LevelIndex = 3, TemplateCode = "04090001", Tentative = true};
        var startNumberingValue4 = new StartNumberingValue {Val = 1};
        var numberingFormat4 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText4 = new LevelText {Val = "·"};
        var levelJustification4 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties4 = new PreviousParagraphProperties();
        var indentation4 = new Indentation {Left = "2880", Hanging = "360"};

        previousParagraphProperties4.Append(indentation4);

        var numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
        var runFonts4 = new RunFonts {Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol"};

        numberingSymbolRunProperties4.Append(runFonts4);

        level4.Append(startNumberingValue4);
        level4.Append(numberingFormat4);
        level4.Append(levelText4);
        level4.Append(levelJustification4);
        level4.Append(previousParagraphProperties4);
        level4.Append(numberingSymbolRunProperties4);

        var level5 = new Level {LevelIndex = 4, TemplateCode = "04090003", Tentative = true};
        var startNumberingValue5 = new StartNumberingValue {Val = 1};
        var numberingFormat5 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText5 = new LevelText {Val = "o"};
        var levelJustification5 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties5 = new PreviousParagraphProperties();
        var indentation5 = new Indentation {Left = "3600", Hanging = "360"};

        previousParagraphProperties5.Append(indentation5);

        var numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
        var runFonts5 = new RunFonts
        {
            Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New",
            ComplexScript = "Courier New"
        };

        numberingSymbolRunProperties5.Append(runFonts5);

        level5.Append(startNumberingValue5);
        level5.Append(numberingFormat5);
        level5.Append(levelText5);
        level5.Append(levelJustification5);
        level5.Append(previousParagraphProperties5);
        level5.Append(numberingSymbolRunProperties5);

        var level6 = new Level {LevelIndex = 5, TemplateCode = "04090005", Tentative = true};
        var startNumberingValue6 = new StartNumberingValue {Val = 1};
        var numberingFormat6 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText6 = new LevelText {Val = "§"};
        var levelJustification6 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties6 = new PreviousParagraphProperties();
        var indentation6 = new Indentation {Left = "4320", Hanging = "360"};

        previousParagraphProperties6.Append(indentation6);

        var numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
        var runFonts6 = new RunFonts {Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings"};

        numberingSymbolRunProperties6.Append(runFonts6);

        level6.Append(startNumberingValue6);
        level6.Append(numberingFormat6);
        level6.Append(levelText6);
        level6.Append(levelJustification6);
        level6.Append(previousParagraphProperties6);
        level6.Append(numberingSymbolRunProperties6);

        var level7 = new Level {LevelIndex = 6, TemplateCode = "04090001", Tentative = true};
        var startNumberingValue7 = new StartNumberingValue {Val = 1};
        var numberingFormat7 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText7 = new LevelText {Val = "·"};
        var levelJustification7 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties7 = new PreviousParagraphProperties();
        var indentation7 = new Indentation {Left = "5040", Hanging = "360"};

        previousParagraphProperties7.Append(indentation7);

        var numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
        var runFonts7 = new RunFonts {Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol"};

        numberingSymbolRunProperties7.Append(runFonts7);

        level7.Append(startNumberingValue7);
        level7.Append(numberingFormat7);
        level7.Append(levelText7);
        level7.Append(levelJustification7);
        level7.Append(previousParagraphProperties7);
        level7.Append(numberingSymbolRunProperties7);

        var level8 = new Level {LevelIndex = 7, TemplateCode = "04090003", Tentative = true};
        var startNumberingValue8 = new StartNumberingValue {Val = 1};
        var numberingFormat8 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText8 = new LevelText {Val = "o"};
        var levelJustification8 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties8 = new PreviousParagraphProperties();
        var indentation8 = new Indentation {Left = "5760", Hanging = "360"};

        previousParagraphProperties8.Append(indentation8);

        var numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
        var runFonts8 = new RunFonts
        {
            Hint = FontTypeHintValues.Default, Ascii = "Courier New", HighAnsi = "Courier New",
            ComplexScript = "Courier New"
        };

        numberingSymbolRunProperties8.Append(runFonts8);

        level8.Append(startNumberingValue8);
        level8.Append(numberingFormat8);
        level8.Append(levelText8);
        level8.Append(levelJustification8);
        level8.Append(previousParagraphProperties8);
        level8.Append(numberingSymbolRunProperties8);

        var level9 = new Level {LevelIndex = 8, TemplateCode = "04090005", Tentative = true};
        var startNumberingValue9 = new StartNumberingValue {Val = 1};
        var numberingFormat9 = new NumberingFormat {Val = NumberFormatValues.Bullet};
        var levelText9 = new LevelText {Val = "§"};
        var levelJustification9 = new LevelJustification {Val = LevelJustificationValues.Left};

        var previousParagraphProperties9 = new PreviousParagraphProperties();
        var indentation9 = new Indentation {Left = "6480", Hanging = "360"};

        previousParagraphProperties9.Append(indentation9);

        var numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
        var runFonts9 = new RunFonts {Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings"};

        numberingSymbolRunProperties9.Append(runFonts9);

        level9.Append(startNumberingValue9);
        level9.Append(numberingFormat9);
        level9.Append(levelText9);
        level9.Append(levelJustification9);
        level9.Append(previousParagraphProperties9);
        level9.Append(numberingSymbolRunProperties9);

        abstractNum1.Append(nsid1);
        abstractNum1.Append(multiLevelType1);
        abstractNum1.Append(templateCode1);
        abstractNum1.Append(level1);
        abstractNum1.Append(level2);
        abstractNum1.Append(level3);
        abstractNum1.Append(level4);
        abstractNum1.Append(level5);
        abstractNum1.Append(level6);
        abstractNum1.Append(level7);
        abstractNum1.Append(level8);
        abstractNum1.Append(level9);
        return abstractNum1;
    }
}