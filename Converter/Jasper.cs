using System;
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
        /// Top-level jasper report document
        /// </summary>
        XmlDocument jDoc;
        /// <summary>
        /// Declaration
        /// </summary>
        XmlDeclaration jDec;
        /// <summary>
        /// Root
        /// </summary>
        XmlElement jRoot;
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
            get { return jDoc.InnerXml; }
        }

        /// <summary>
        /// Initialise the XML objects
        /// </summary>
        protected void InitXML()
        {
            jDoc = new XmlDocument();
            jDec = jDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            jRoot = jDoc.DocumentElement;
            jDoc.InsertBefore(jDec, jRoot);

            XmlComment comment = jDoc.CreateComment("Created by Word Report Conversion by Wassey Development");
            jDoc.InsertBefore(comment, jRoot);

            // Main jasperReport element - set the page attributes based upon the document
            jJasper = jDoc.CreateElement("jasperReport");
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
            jJasper.SetAttribute("pageWidth", ((int)WordDoc.PageSetup.PageWidth).ToString());
            jJasper.SetAttribute("pageHeight", ((int)WordDoc.PageSetup.PageHeight).ToString());
            jJasper.SetAttribute("leftMargin", ((int)WordDoc.PageSetup.LeftMargin).ToString());
            jJasper.SetAttribute("rightMargin", ((int)WordDoc.PageSetup.RightMargin).ToString());
            jJasper.SetAttribute("topMargin", ((int)WordDoc.PageSetup.TopMargin).ToString());
            jJasper.SetAttribute("bottomMargin", ((int)WordDoc.PageSetup.BottomMargin).ToString());

            // Other miscellaneous attributes
            jJasper.SetAttribute("name", WordDoc.Name);
            jJasper.SetAttribute("whenNoDataType", "NoDataSection");

            // Set the current band type to detail by default
            currentBandType = BandType.Detail;

        }
    }
}
