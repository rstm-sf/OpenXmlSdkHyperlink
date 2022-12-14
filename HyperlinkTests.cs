using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

using Color = SixLabors.ImageSharp.Color;

namespace OpenXmlSdkHyperlink;

[UsesVerify]
public class HyperlinkTests
{
    [Fact]
public async Task AdvancedWordCreate() {
        using var document = WordDocument.Create();
        // lets add some properties to the document
        document.BuiltinDocumentProperties.Title = "Cover Page Templates";
        document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";
        document.ApplicationProperties.Company = "Evotec Services";

        // we force document to update fields on open, this will be used by TOC
        document.Settings.UpdateFieldsOnOpen = true;

        // lets add one of multiple added Cover Pages
        document.AddCoverPage(CoverPageTemplate.IonDark);

        // lets add Table of Content (1 of 2)
        document.AddTableOfContent(TableOfContentStyle.Template1);

        // lets add page break
        document.AddPageBreak();

        // lets create a list that will be binded to TOC
        var wordListToc = document.AddTableOfContentList(WordListStyle.Headings111);

        wordListToc.AddItem("How to add a table to document?");

        document.AddParagraph(
            "In the first paragraph I would like to show you how to add a table to the document using one of the 105 built-in styles:");

        // adding a table and modifying content
        var table = document.AddTable(5, 4, WordTableStyle.GridTable5DarkAccent5);
        table.Rows[3].Cells[2].Paragraphs[0].Text = "Adding text to cell";
        table.Rows[3].Cells[2].Paragraphs[0].Color = Color.Blue;
        ;
        table.Rows[3].Cells[3].Paragraphs[0].Text = "Different cell";

        document.AddParagraph("As you can see adding a table with some style, and adding content to it ").SetBold()
            .SetUnderline(UnderlineValues.Dotted).AddText("is not really complicated").SetColor(Color.OrangeRed);

        wordListToc.AddItem("How to add a list to document?");

        var paragraph = document
            .AddParagraph("Adding lists is similar to adding a table. Just define a list and add list items to it. ")
            .SetText("Remember that you can add anything between list items! ");
        paragraph.SetColor(Color.Blue)
            .SetText("For example TOC List is just another list, but defining a specific style.");

        var list = document.AddList(WordListStyle.Bulleted);
        list.AddItem("First element of list", 0);
        list.AddItem("Second element of list", 1);

        var paragraphWithHyperlink = document.AddHyperLink("Go to Evotec Blogs", new Uri("https://evotec.xyz"),
            true, "URL with tooltip");
        // you can also change the hyperlink text, uri later on using properties
        paragraphWithHyperlink.Hyperlink.Uri = new Uri("https://evotec.xyz/hub");
        paragraphWithHyperlink.ParagraphAlignment = JustificationValues.Center;

        list.AddItem("3rd element of list, but added after hyperlink", 0);
        list.AddItem("4th element with hyperlink ")
            .AddHyperLink("included.", new Uri("https://evotec.xyz/hub"), addStyle: true);

        document.AddParagraph();

        var listNumbered = document.AddList(WordListStyle.Heading1ai);
        listNumbered.AddItem("Different list number 1");
        listNumbered.AddItem("Different list number 2", 1);
        listNumbered.AddItem("Different list number 3", 1);
        listNumbered.AddItem("Different list number 4", 1);

        var section = document.AddSection();
        section.PageOrientation = PageOrientationValues.Landscape;
        section.PageSettings.PageSize = WordPageSize.A4;

        wordListToc.AddItem("Adding headers / footers");

        // lets add headers and footers
        document.AddHeadersAndFooters();

        // adding text to default header
        document.Header.Default.AddParagraph("Text added to header - Default");

        var section1 = document.AddSection();
        section1.PageOrientation = PageOrientationValues.Portrait;
        section1.PageSettings.PageSize = WordPageSize.A5;

        wordListToc.AddItem("Adding custom properties and page numbers to document");

        document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty {Value = DateTime.Today});
        document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
        document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

        // add page numbers
        document.Footer.Default.AddPageNumber(WordPageNumberStyle.PlainNumber);

        // add watermark
        document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Draft");

        document.Save();

        var i = 1;
        foreach (var hyperlink in document._document.Descendants<Hyperlink>()) {
            hyperlink.Id = "R" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var headerReference in document._document.Descendants<HeaderReference>()) {
            headerReference.Id = "R" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var footerReference in document._document.Descendants<FooterReference>()) {
            footerReference.Id = "R" + i.ToString("X8");
            i++;
        }

        await Verifier.Verify(FormatXml(document._document.OuterXml));
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
        var levelText1 = new LevelText {Val = "??"};
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
        var levelText3 = new LevelText {Val = "??"};
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
        var levelText4 = new LevelText {Val = "??"};
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
        var levelText6 = new LevelText {Val = "??"};
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
        var levelText7 = new LevelText {Val = "??"};
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
        var levelText9 = new LevelText {Val = "??"};
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

    private static void SaveNumbering(WordprocessingDocument docPackage)
    {
        var numbering = docPackage.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;

        // it seems the order of numbering instance/abstractnums in numbering matters...

        var listAbstractNum = numbering.ChildElements.OfType<AbstractNum>().ToArray();
        var listNumberingInstance = numbering.ChildElements.OfType<NumberingInstance>().ToArray();
        var listNumberPictures = numbering.ChildElements.OfType<NumberingPictureBullet>().ToArray();

        numbering.RemoveAllChildren();

        foreach (var pictureBullet in listNumberPictures)
            numbering.Append(pictureBullet);

        foreach (var abstractNum in listAbstractNum)
            numbering.Append(abstractNum);

        foreach (var numberingInstance in listNumberingInstance)
            numbering.Append(numberingInstance);
    }

    private void MoveSectionProperties(WordprocessingDocument docPackage)
    {
        var body = docPackage.MainDocumentPart!.Document.Body!;
        var sectionProperties = body.Elements<SectionProperties>().Last();
        body.RemoveChild(sectionProperties);
        body.Append(sectionProperties);
    }
}