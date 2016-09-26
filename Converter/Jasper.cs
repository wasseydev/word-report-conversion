﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Xml;

namespace Converter
{
    public enum BandType
    {
        none,
        Background,
        Title,
        PageHeader,
        ColumnHeader,
        Detail,
        ColumnFooter,
        PageFooter,
        LastPageFooter,
        Summary,
        NoData
    }

    public class Jasper
    {
        /// <summary>
        /// The Word Application containing the documents that are being converted.
        /// </summary>
        Word.Application WordApp;
        /// <summary>
        /// If only a single document is being converted, then this is it.
        /// </summary>
        Word.Document WordDoc;
        /// <summary>
        /// Page setup of the document
        /// </summary>
        Word.PageSetup pageSetup;
        /// <summary>
        /// Column width of the document (or current section possibly)?
        /// </summary>
        float columnWidth;
        /// <summary>
        /// Top-level jasper report document
        /// </summary>
        XmlDocument jDoc;
        /// <summary>
        /// Declaration
        /// </summary>
        XmlDeclaration jDec;
        /// <summary>
        /// Jasper element
        /// </summary>
        XmlElement jJasper;
        /// <summary>
        /// Query string for the report document
        /// </summary>
        XmlElement jQuery;
        /// <summary>
        /// Background element
        /// </summary>
        XmlElement jBackground;
        /// <summary>
        /// Title element
        /// </summary>
        XmlElement jTitle;
        /// <summary>
        /// Page header element
        /// </summary>
        XmlElement jPageHeader;
        /// <summary>
        /// Column header element
        /// </summary>
        XmlElement jColumnHeader;
        /// <summary>
        /// Detail element
        /// </summary>
        XmlElement jDetail;
        /// <summary>
        /// Column footer element
        /// </summary>
        XmlElement jColumnFooter;
        /// <summary>
        /// Page footer element
        /// </summary>
        XmlElement jPageFooter;
        /// <summary>
        /// Last page footer element
        /// </summary>
        XmlElement jLastPageFooter;
        /// <summary>
        /// Summary element
        /// </summary>
        XmlElement jSummary;
        /// <summary>
        /// No data element
        /// </summary>
        XmlElement jNoData;
        /// <summary>
        /// The current band type that we are processing
        /// </summary>
        BandType currentBandType;

        /// <summary>
        /// Constructor that takes a Word.Application reference. In this case, the
        /// document is considered to be the active document. If there is no active
        /// document, an exception is automatically thrown.
        /// </summary>
        /// <param name="wordApp">Word.Application</param>
        public Jasper(Word.Application wordApp)
        {
            WordApp = wordApp;
            WordDoc = WordApp.ActiveDocument;
            SetDefaultPreferences();
            InitXML();
        }

        /// <summary>
        /// Constructor that takes a Word.Application reference and a Document. If either
        /// parameter is null, an exception is thrown.
        /// </summary>
        /// <param name="wordApp">Word.Application</param>
        /// <param name="wordDoc">Word.Document</param>
        public Jasper(Word.Application wordApp, Word.Document wordDoc)
        {
            WordApp = wordApp;
            WordDoc = wordDoc;
            if (WordApp == null)
            {
                throw new Exception("Word Application cannot be null");
            }
            else if (WordDoc == null)
            {
                throw new Exception("No active document");
            }
            SetDefaultPreferences();
            InitXML();
        }

        /// <summary>
        /// Set the default preferences for Jasper conversion
        /// </summary>
        public void SetDefaultPreferences()
        {
            //TODO: Complete this
        }

        public String JRXML {
            get {
                //TODO: Properly fix prefix on xsi:schemaLocation
                String x = jDoc.InnerXml;
                x = x.Replace("d1p1:schemaLocation=\"", "xsi:schemaLocation=\"");
                return x;
            }
        }

        /// <summary>
        /// Initialise the XML objects
        /// </summary>
        protected void InitXML()
        {
            jDoc = new XmlDocument();
            
            jDec = jDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            jDoc.AppendChild(jDec);

            XmlComment comment = jDoc.CreateComment("Created by Word Report Conversion by Wassey Development");
            jDoc.AppendChild(comment);

            // Main jasperReport element
            jJasper = jDoc.CreateElement("jasperReport");
            jJasper.SetAttribute("xmlns", "http://jasperreports.sourceforge.net/jasperreports");
            jJasper.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            jJasper.SetAttribute("schemaLocation", "http://jasperreports.sourceforge.net/jasperreports", "http://jasperreports.sourceforge.net/jasperreports http://jasperreports.sourceforge.net/xsd/jasperreport.xsd");
            jDoc.AppendChild(jJasper);

            jQuery = jDoc.CreateElement("queryString");
            jJasper.AppendChild(jQuery);

            jBackground = jDoc.CreateElement("background");
            jJasper.AppendChild(jBackground);

            jTitle = jDoc.CreateElement("title");
            jJasper.AppendChild(jTitle);

            jPageHeader = jDoc.CreateElement("pageHeader");
            jJasper.AppendChild(jPageHeader);

            jColumnHeader = jDoc.CreateElement("columnHeader");
            jJasper.AppendChild(jColumnHeader);

            jDetail = jDoc.CreateElement("detail");
            jJasper.AppendChild(jDetail);

            jColumnFooter = jDoc.CreateElement("columnFooter");
            jJasper.AppendChild(jColumnFooter);

            jPageFooter = jDoc.CreateElement("pageFooter");
            jJasper.AppendChild(jPageFooter);

            jLastPageFooter = jDoc.CreateElement("lastPageFooter");
            jJasper.AppendChild(jLastPageFooter);

            jSummary = jDoc.CreateElement("summary");
            jJasper.AppendChild(jSummary);

            jNoData = jDoc.CreateElement("noData");
            jJasper.AppendChild(jNoData);

            Debug.WriteLine(jDoc.InnerXml);
        }

        /// <summary>
        /// Process the document, converting it into JRXML
        /// </summary>
        public void ProcessDocument()
        {
            // Complete the layout attributes
            pageSetup = WordDoc.PageSetup;
            columnWidth = pageSetup.PageWidth - pageSetup.RightMargin - pageSetup.LeftMargin;

            jJasper.SetAttribute("pageWidth", ((int)pageSetup.PageWidth).ToString());
            jJasper.SetAttribute("pageHeight", ((int)pageSetup.PageHeight).ToString());
            jJasper.SetAttribute("leftMargin", ((int)pageSetup.LeftMargin).ToString());
            jJasper.SetAttribute("rightMargin", ((int)pageSetup.RightMargin).ToString());
            jJasper.SetAttribute("topMargin", ((int)pageSetup.TopMargin).ToString());
            jJasper.SetAttribute("bottomMargin", ((int)pageSetup.BottomMargin).ToString());
            jJasper.SetAttribute("columnWidth", ((int)columnWidth).ToString());
            
            // Other miscellaneous attributes
            jJasper.SetAttribute("name", WordDoc.Name);
            jJasper.SetAttribute("whenNoDataType", "NoDataSection");

            // Set the current band type to detail by default
            currentBandType = BandType.Detail;

            // Iterate through the paragraphs
            foreach (Word.Paragraph paragraph in WordDoc.Paragraphs)
            {
                ProcessParagraph(paragraph);
            }
        }

        protected XmlElement GetBandElement(BandType bandType)
        {
            switch (bandType)
            {
                case BandType.Background:
                    return jBackground;
                case BandType.Title:
                    return jTitle;
                case BandType.PageHeader:
                    return jPageHeader;
                case BandType.ColumnHeader:
                    return jColumnHeader;
                case BandType.Detail:
                    return jDetail;
                case BandType.ColumnFooter:
                    return jColumnFooter;
                case BandType.PageFooter:
                    return jPageFooter;
                case BandType.LastPageFooter:
                    return jLastPageFooter;
                case BandType.Summary:
                    return jSummary;
                case BandType.NoData:
                    return jNoData;
                default:
                    return null;
            }
        }

        /// <summary>
        /// Parse the provided tag text to ensure it is valid and return the tag
        /// and inner tag value. The text supplied should start with the $ character,
        /// and can have any text following the closing }, but must include both
        /// { and }. There is no specific validation of tag names and values.
        /// </summary>
        /// <param name="tagText">The text for the tag</param>
        /// <param name="tag">The tag from the text i.e. what was after the $.
        /// It is changed to lower case.</param>
        /// <param name="inner">The inner value of the tag text, i.e. what is 
        /// between {}</param>
        /// <param name="nextChar">The position of the next character in tagText
        /// following the closing brace</param>
        /// <returns>true if the tag text was valid, false otherwise</returns>
        protected bool ParseTagText(String tagText, out String tag, out String inner,
            out int nextChar)
        {
            tag = String.Empty;
            inner = String.Empty;
            nextChar = -1;

            // Must start with $
            if (!tagText.StartsWith("$")) return false;

            // Tag name is between $ and {, with no } between.
            int openPos = tagText.IndexOf('{');
            int closePos = tagText.IndexOf('}');

            // Invalid if close is before open or if either is not present
            if (closePos < openPos || closePos < 0 || openPos < 0) return false;

            tag = tagText.Substring(1, openPos - 1).ToLower();
            inner = tagText.Substring(openPos + 1, closePos - openPos - 1);
            nextChar = closePos + 1;
            return true;
        }

        public void ProcessParagraph(Word.Paragraph paragraph)
        {
            XmlElement band;
            String text = paragraph.Range.Text;
            text = text.Substring(0, text.Length - 1);
            Word.Style style = paragraph.get_Style();
            Word.ParagraphFormat paraFormat = paragraph.Format;
            
            int spaceBefore = (int)paraFormat.SpaceBefore;
            int spaceAfter = (int)paraFormat.SpaceAfter;
            int fontSize = (int)style.Font.Size;
            int bandHeight = spaceBefore + spaceAfter + fontSize;

            Debug.WriteLine("Processing paragraph: " + text);
            int parseFromChar = 0;

            if (text.StartsWith("$"))
            {
                String tagName, tagValue;
                int nextChar;

                if (ParseTagText(text, out tagName, out tagValue, out nextChar))
                {
                    if (tagName.Equals("bandtype"))
                    {
                        switch (tagValue.ToLower())
                        {
                            case "background":
                                currentBandType = BandType.Background;
                                break;
                            case "title":
                                currentBandType = BandType.Title;
                                break;
                            case "pageheader":
                                currentBandType = BandType.PageHeader;
                                break;
                            case "columnheader":
                                currentBandType = BandType.ColumnHeader;
                                break;
                            case "detail":
                                currentBandType = BandType.Detail;
                                break;
                            case "columnfooter":
                                currentBandType = BandType.ColumnFooter;
                                break;
                            case "pagefooter":
                                currentBandType = BandType.PageFooter;
                                break;
                            case "lastpagefooter":
                                currentBandType = BandType.LastPageFooter;
                                break;
                            case "summary":
                                currentBandType = BandType.Summary;
                                break;
                            case "nodata":
                                currentBandType = BandType.NoData;
                                break;
                            default:
                                Debug.WriteLine("Invalid band type " + tagValue);
                                break;
                        }
                        Debug.WriteLine("Changed band type to " + currentBandType.ToString());
                        return;
                    }
                }
            }
            if (currentBandType == BandType.Detail)
            {
                // Add a new band for each paragraph
                band = jDoc.CreateElement("band");
                jDetail.AppendChild(band);

                // Set the band attributes
                // Set the height based upon the paragraph style - it will stretch

                band.SetAttribute("height", bandHeight.ToString());
                // Split type
                band.SetAttribute("splitType", "Stretch");
            }
            else
            {
                // There is only one band of these types, so all paragraphs have to fit
                XmlElement bandElement = GetBandElement(currentBandType);
                
                // If there is a child element, it will be the band. If not, create it
                if (bandElement.HasChildNodes)
                {
                    XmlNode node = bandElement.GetElementsByTagName("band")[0];
                    band = (XmlElement)node;
                }
                else
                {
                    // Add a new band covering all paragraphs
                    band = jDoc.CreateElement("band");
                    bandElement.AppendChild(band);
                    band.SetAttribute("height", bandHeight.ToString());
                    band.SetAttribute("splitType", "Stretch");
                }
            }

            // Add the text field, report element, text element, and text field expression
            XmlElement textField = jDoc.CreateElement("textField");
            XmlElement reportElt = jDoc.CreateElement("reportElement");
            XmlElement textElt = jDoc.CreateElement("textElement");
            XmlElement textFont = jDoc.CreateElement("font");
            XmlElement textFieldExp = jDoc.CreateElement("textFieldExpression");

            band.AppendChild(textField);
            textField.AppendChild(reportElt);
            textField.AppendChild(textElt);
            textField.AppendChild(textFieldExp);
            textElt.AppendChild(textFont);

            textField.SetAttribute("isStretchWithOverflow", "true");
            textField.SetAttribute("isBlankWhenNull", "true");

            reportElt.SetAttribute("stretchType", "RelativeToBandHeight");
            // TODO: Set these attributes to be dynamic based upon paragraph
            reportElt.SetAttribute("x", "0");
            reportElt.SetAttribute("width", ((int)columnWidth).ToString());
            // TODO: If not detail band, then y and height need to be calculated
            // TODO: y position should consider the paragraph spacing before
            reportElt.SetAttribute("y", spaceBefore.ToString());
            reportElt.SetAttribute("height", fontSize.ToString());

            textElt.SetAttribute("markup", "styled");
            
            // Get the format of the first character we are parsing. This becomes
            // our 'base' format.
            int textLength = text.Length;
            Word.Characters chars = paragraph.Range.Characters;

            Word.Range refRange = chars[parseFromChar + 1];

            textFont.SetAttribute("fontName", refRange.Font.Name);
            textFont.SetAttribute("size", refRange.Font.Size.ToString());

            Word.Range lastRange = refRange;

            String openStyle, closeStyle;
            GetStyleTags(refRange, refRange, out openStyle, out closeStyle);
            StringBuilder jText = new StringBuilder(openStyle, 2048);
            
            for (int c = parseFromChar; c < textLength; c++)
            {
                if (c > parseFromChar)
                {
                    // Compare the style to refRange, and if style is different,
                    // close previousstyle tag and create a new one
                    Word.Range curRange = chars[c + 1];
                    if (!AreStylesEqual(lastRange, curRange))
                    {
                        // Close the previous style
                        jText.Append(closeStyle);
                        // Start the new style
                        GetStyleTags(curRange, refRange, out openStyle, out closeStyle);
                        jText.Append(openStyle);
                        lastRange = curRange;
                    }
                }
                // TODO: Cater for fields, etc
                String ch = text.Substring(c, 1);
                if (ch == "<")
                {
                    ch = "&lt;";
                }
                else if (ch == ">")
                {
                    ch = "&gt;";
                }
                else if (ch == "&")
                {
                    ch = "&amp;";
                }
                else if (ch == "\"")
                {
                    ch = "&quot;";
                }
                else if (ch == "'")
                {
                    ch = "&apos;";
                }
                else if (ch == "\r" || ch == "\n" || ch == "\v")
                {
                    ch = "<br/>";
                }
                jText.Append(ch);
            }
            // Close the style
            jText.Append(closeStyle);

            XmlCDataSection styledText = jDoc.CreateCDataSection("\"" + jText + "\"");
            textFieldExp.AppendChild(styledText);
        }

        protected bool AreStylesEqual(Word.Range r1, Word.Range r2)
        {
            if (
                r1.Bold != r2.Bold
                || r1.Italic != r2.Italic
                || (r1.Underline == Word.WdUnderline.wdUnderlineNone)
                    != (r2.Underline == Word.WdUnderline.wdUnderlineNone)
                )
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Get the style tag for the current range. It will consider attributes
        /// such as bold, italic, etc, and if these are set, it will return the
        /// appropriate open and close tags. If there are no attributes set, then
        /// no tags will be returned.<br/>
        /// Attributes including the font name and size will be compared against
        /// the reference range, and if they differ, attributes will be set.
        /// </summary>
        /// <param name="curRange">The current range to be considered</param>
        /// <param name="refRange">The reference range to be compared against</param>
        /// <param name="openTag">Output of the opening style tag</param>
        /// <param name="closeTag">Output of the closing style tag</param>
        protected void GetStyleTags(Word.Range curRange, Word.Range refRange,
            out String openTag, out String closeTag)
        {
            bool styled = false;
            openTag = "";
            closeTag = "";

            StringBuilder tagText = new StringBuilder("<style", 50);
            if (curRange.Bold != 0)
            {
                tagText.Append(" isBold=\\\"true\\\"");
                styled = true;
            }
            if (curRange.Italic != 0)
            {
                tagText.Append(" isItalic=\\\"true\\\"");
                styled = true;
            }
            if (curRange.Underline != Word.WdUnderline.wdUnderlineNone)
            {
                tagText.Append(" isUnderline=\\\"true\\\"");
                styled = true;
            }
            tagText.Append(">");

            // Only update the output style tags if there was an attribute set
            if (styled)
            {
                openTag = tagText.ToString();
                closeTag = "</style>";
            }
        }
    }
}
