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
        public void InitXML()
        {
            jDoc = new XmlDocument();
            jDec = jDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            jRoot = jDoc.DocumentElement;
            jDoc.InsertBefore(jDec, jRoot);

            XmlComment comment = jDoc.CreateComment("Created by Word Report Conversion by Wassey Development");
            jDoc.InsertBefore(comment, jRoot);

            // Main jasperReport element - set the page attributes based upon the document
            jJasper = jDoc.CreateElement("jasperReport");
            jJasper.SetAttribute("name", WordDoc.Name);
            jJasper.SetAttribute("pageWidth", ((int) WordDoc.PageSetup.PageWidth).ToString());
            jJasper.SetAttribute("pageHeight", ((int) WordDoc.PageSetup.PageHeight).ToString());
            jJasper.SetAttribute("leftMargin", ((int) WordDoc.PageSetup.LeftMargin).ToString());
            jJasper.SetAttribute("rightMargin", ((int) WordDoc.PageSetup.RightMargin).ToString());
            jJasper.SetAttribute("topMargin", ((int) WordDoc.PageSetup.TopMargin).ToString());
            jJasper.SetAttribute("bottomMargin", ((int) WordDoc.PageSetup.BottomMargin).ToString());

            jDoc.AppendChild(jJasper);

            jQuery = jDoc.CreateElement("queryString");
            jJasper.AppendChild(jQuery);



            Debug.WriteLine(jDoc.InnerXml);
        }
    }
}
