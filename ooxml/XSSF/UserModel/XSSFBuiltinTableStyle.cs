/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using EnumsNET;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace NPOI.XSSF.UserModel
{
    /// <summary>
    /// Table style names defined in the OOXML spec.
    /// The actual styling is defined in presetTableStyles.xml
    /// </summary>
    public enum XSSFBuiltinTableStyleEnum : int
    {
        TableStyleDark1,

        TableStyleDark2,

        TableStyleDark3,

        TableStyleDark4,

        TableStyleDark5,

        TableStyleDark6,

        TableStyleDark7,

        TableStyleDark8,

        TableStyleDark9,

        TableStyleDark10,

        TableStyleDark11,

        TableStyleLight1,

        TableStyleLight2,

        TableStyleLight3,

        TableStyleLight4,

        TableStyleLight5,

        TableStyleLight6,

        TableStyleLight7,

        TableStyleLight8,

        TableStyleLight9,

        TableStyleLight10,

        TableStyleLight11,

        TableStyleLight12,

        TableStyleLight13,

        TableStyleLight14,

        TableStyleLight15,

        TableStyleLight16,

        TableStyleLight17,

        TableStyleLight18,

        TableStyleLight19,

        TableStyleLight20,

        TableStyleLight21,

        TableStyleMedium1,

        TableStyleMedium2,

        TableStyleMedium3,

        TableStyleMedium4,

        TableStyleMedium5,

        TableStyleMedium6,

        TableStyleMedium7,

        TableStyleMedium8,

        TableStyleMedium9,

        TableStyleMedium10,

        TableStyleMedium11,

        TableStyleMedium12,

        TableStyleMedium13,

        TableStyleMedium14,

        TableStyleMedium15,

        TableStyleMedium16,

        TableStyleMedium17,

        TableStyleMedium18,

        TableStyleMedium19,

        TableStyleMedium20,

        TableStyleMedium21,

        TableStyleMedium22,

        TableStyleMedium23,

        TableStyleMedium24,

        TableStyleMedium25,

        TableStyleMedium26,

        TableStyleMedium27,

        TableStyleMedium28,

        PivotStyleMedium1,

        PivotStyleMedium2,

        PivotStyleMedium3,

        PivotStyleMedium4,

        PivotStyleMedium5,

        PivotStyleMedium6,

        PivotStyleMedium7,

        PivotStyleMedium8,

        PivotStyleMedium9,

        PivotStyleMedium10,

        PivotStyleMedium11,

        PivotStyleMedium12,

        PivotStyleMedium13,

        PivotStyleMedium14,

        PivotStyleMedium15,

        PivotStyleMedium16,

        PivotStyleMedium17,

        PivotStyleMedium18,

        PivotStyleMedium19,

        PivotStyleMedium20,

        PivotStyleMedium21,

        PivotStyleMedium22,

        PivotStyleMedium23,

        PivotStyleMedium24,

        PivotStyleMedium25,

        PivotStyleMedium26,

        PivotStyleMedium27,

        PivotStyleMedium28,

        PivotStyleLight1,

        PivotStyleLight2,

        PivotStyleLight3,

        PivotStyleLight4,

        PivotStyleLight5,

        PivotStyleLight6,

        PivotStyleLight7,

        PivotStyleLight8,

        PivotStyleLight9,

        PivotStyleLight10,

        PivotStyleLight11,

        PivotStyleLight12,

        PivotStyleLight13,

        PivotStyleLight14,

        PivotStyleLight15,

        PivotStyleLight16,

        PivotStyleLight17,

        PivotStyleLight18,

        PivotStyleLight19,

        PivotStyleLight20,

        PivotStyleLight21,

        PivotStyleLight22,

        PivotStyleLight23,

        PivotStyleLight24,

        PivotStyleLight25,

        PivotStyleLight26,

        PivotStyleLight27,

        PivotStyleLight28,

        PivotStyleDark1,

        PivotStyleDark2,

        PivotStyleDark3,

        PivotStyleDark4,

        PivotStyleDark5,

        PivotStyleDark6,

        PivotStyleDark7,

        PivotStyleDark8,

        PivotStyleDark9,

        PivotStyleDark10,

        PivotStyleDark11,

        PivotStyleDark12,

        PivotStyleDark13,

        PivotStyleDark14,

        PivotStyleDark15,

        PivotStyleDark16,

        PivotStyleDark17,

        PivotStyleDark18,

        PivotStyleDark19,

        PivotStyleDark20,

        PivotStyleDark21,

        PivotStyleDark22,

        PivotStyleDark23,

        PivotStyleDark24,

        PivotStyleDark25,

        PivotStyleDark26,

        PivotStyleDark27,

        PivotStyleDark28
    }
    public class XSSFBuiltinTableStyle
    {
        const string presetTableStylesResourceName = "NPOI.Resources.presetTableStyles.xml";

        private static Dictionary<XSSFBuiltinTableStyleEnum, ITableStyle> styleMap = new Dictionary<XSSFBuiltinTableStyleEnum, ITableStyle>();

        /// <summary>
        /// Experimental switch: when true, use XDocument-based loader in Init().
        /// Default is false to preserve legacy behavior.
        /// </summary>
        public static bool UseXDocument { get; set; } =
            string.Equals(Environment.GetEnvironmentVariable("NPOI_XSSF_BUILTIN_TABLESTYLE_USE_XDOCUMENT"), "1", StringComparison.OrdinalIgnoreCase);

        // ---- benchmark/testing helpers (keep internal) ----
        internal static void ResetForTesting()
        {
            styleMap.Clear();
        }

        internal static void EnsureInitializedForTesting()
        {
            Init();
        }

        public static ITableStyle GetStyle(XSSFBuiltinTableStyleEnum style)
        {
            Init();
            return styleMap[style];
        }

        public static bool IsBuiltinStyle(ITableStyle style)
        {
            if (style == null)
                return false;
            return Enums.GetNames<XSSFBuiltinTableStyleEnum>().Any(x => x == style.Name);
        }

        private static void Init()
        {
            if (styleMap.Count > 0)
                return;

            if (UseXDocument)
            {
                InitViaXDocument();
            }
            else
            {
                InitViaXmlDocument();
            }
        }

        /// <summary>
        /// Legacy implementation (existing code). Keep as-is as baseline.
        /// </summary>
        private static void InitViaXmlDocument()
        {
            using (var xmlstream = typeof(XSSFBuiltinTableStyle).Assembly.GetManifestResourceStream(presetTableStylesResourceName))
            {
                var xmlReader = new XmlTextReader(xmlstream);
                var xmlDocument = new XmlDocument();
                xmlDocument.Load(xmlReader);

                var node = xmlDocument.SelectSingleNode("presetTableStyles");
                foreach (XmlNode child in node.ChildNodes)
                {
                    String styleName = child.Name;
                    if (!Enum.TryParse<XSSFBuiltinTableStyleEnum>(styleName, out XSSFBuiltinTableStyleEnum builtIn))
                    {
                        continue;
                    }
                    var dxfsNode = child.SelectSingleNode("dxfs");
                    var tableStyleNode = child.SelectSingleNode("tableStyles");
                    StylesTable styles = new StylesTable();
                    var styleXmlDocument = new XmlDocument();
                    styleXmlDocument.LoadXml(StyleXML(dxfsNode, tableStyleNode));
                    styles.ReadFrom(styleXmlDocument);
                    styleMap.Add(builtIn, new XSSFBuiltinTypeStyleStyle(builtIn, styles.GetExplicitTableStyle(styleName)));
                }
            }
        }

        /// <summary>
        /// Experimental XDocument implementation.
        /// Produces the same styleSheet XML string, then feeds it into existing StylesTable.ReadFrom(XmlDocument)
        /// to keep downstream behavior unchanged.
        /// </summary>
        private static void InitViaXDocument()
        {
            using (var xmlstream = typeof(XSSFBuiltinTableStyle).Assembly.GetManifestResourceStream(presetTableStylesResourceName))
            {
                if (xmlstream == null)
                    throw new InvalidOperationException($"Embedded resource not found: {presetTableStylesResourceName}");

                // Use a safe XmlReader (avoid DTD/entity resolution)
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Ignore,
                    XmlResolver = null
                };

                using var reader = XmlReader.Create(xmlstream, settings);
                var xdoc = XDocument.Load(reader, LoadOptions.None);

                var root = xdoc.Root;
                if (root == null || root.Name.LocalName != "presetTableStyles")
                    throw new InvalidOperationException("presetTableStyles root element not found.");

                foreach (var styleElem in root.Elements())
                {
                    // In this resource there are no namespaces/prefixes, LocalName is safest anyway
                    var styleName = styleElem.Name.LocalName;

                    if (!Enum.TryParse<XSSFBuiltinTableStyleEnum>(styleName, out var builtIn))
                        continue;

                    var dxfsElem = styleElem.Element("dxfs");
                    var tableStylesElem = styleElem.Element("tableStyles");

                    if (dxfsElem == null || tableStylesElem == null)
                        continue;

                    StylesTable styles = new StylesTable();
                    var styleXmlDocument = new XmlDocument();

                    // Build styleSheet XML
                    var styleXml = StyleXML(dxfsElem, tableStylesElem);

                    styleXmlDocument.LoadXml(styleXml);
                    styles.ReadFrom(styleXmlDocument);

                    styleMap.Add(builtIn, new XSSFBuiltinTypeStyleStyle(builtIn, styles.GetExplicitTableStyle(styleName)));
                }
            }
        }

        // ---- existing StyleXML(XmlNode, XmlNode) remains unchanged ----
        private static string StyleXML(XmlNode dxfsNode, XmlNode tableStyleNode)
        {
            // built-ins doc uses 1-based dxf indexing, Excel uses 0 based.
            // add a dummy node to adjust properly.
            dxfsNode.InsertBefore(dxfsNode.OwnerDocument.CreateElement("dxf"), dxfsNode.FirstChild);

            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
                    .Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ")
                    .Append("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ")
                    .Append("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" ")
                    .Append("xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" ")
                    .Append("mc:Ignorable=\"x14ac x16r2\">\n");
            sb.Append(dxfsNode.OuterXml);
            sb.Append(tableStyleNode.OuterXml);
            sb.Append("</styleSheet>");
            return sb.ToString();
        }

        // ---- new overload for XDocument path ----
        private static string StyleXML(XElement dxfsElem, XElement tableStyleElem)
        {
            // built-ins doc uses 1-based dxf indexing, Excel uses 0 based.
            // add a dummy node to adjust properly.
            // No namespaces in the resource, so "dxf" is fine.
            dxfsElem.AddFirst(new XElement("dxf"));

            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
                    .Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ")
                    .Append("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ")
                    .Append("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" ")
                    .Append("xmlns:x16r2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main\" ")
                    .Append("mc:Ignorable=\"x14ac x16r2\">\n");

            // OuterXml equivalent (no indentation) for stability/perf
            sb.Append(dxfsElem.ToString(SaveOptions.DisableFormatting));
            sb.Append(tableStyleElem.ToString(SaveOptions.DisableFormatting));

            sb.Append("</styleSheet>");
            return sb.ToString();
        }
        protected class XSSFBuiltinTypeStyleStyle : ITableStyle
        {

            private XSSFBuiltinTableStyleEnum builtIn;
            private ITableStyle style;

            /**
             * @param builtIn
             * @param style
             */
            internal XSSFBuiltinTypeStyleStyle(XSSFBuiltinTableStyleEnum builtIn, ITableStyle style)
            {
                this.builtIn = builtIn;
                this.style = style;
            }

            public String Name
            {
                get
                {
                    return style.Name;
                }
            }

            public int Index
            {
                get
                {
                    return (int) builtIn;
                }
            }

            public bool IsBuiltin
            {
                get
                {
                    return true;
                }
            }

            public IDifferentialStyleProvider GetStyle(TableStyleType type)
            {
                return style.GetStyle(type);
            }

        }
    }
}
