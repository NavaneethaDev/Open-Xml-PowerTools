/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    internal class WmlRunSplitter
    {
        internal static void Split(WordprocessingDocument wDoc)
        {
            SplitAllElements(wDoc);
        }

        internal static XDocument Coalesce(SplitRunsAnnotation splitRunsAnnotation)
        {
            XDocument newXDoc = new XDocument();
            var newBodyChildren = CoalesceRecurse(splitRunsAnnotation.SplitRuns, 0);
            newXDoc.Add(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.pt.NamespaceName),
                new XElement(W.body, newBodyChildren)));
            return newXDoc;
        }

        private static object CoalesceRecurse(IEnumerable<SplitRun> list, int level)
        {
            var grouped = list
                .GroupBy(sr =>
                {
                    if (level >= sr.AncestorElements.Length)
                        return Guid.NewGuid().ToString().Replace("-", "");
                    var unid = (string)sr.AncestorElements[level].Attribute(PtOpenXml.Unid);
                    return unid;
                });

#if false
            var sb = new StringBuilder();
            foreach (var group in grouped)
            {
                sb.AppendFormat("Group Key: {0}", group.Key);
                sb.Append(Environment.NewLine);
                foreach (var groupChildItem in group)
                {
                    sb.Append("  ");
                    sb.Append(groupChildItem.ToString(0));
                    sb.Append(Environment.NewLine);
                }
                sb.Append(Environment.NewLine);
            }
#endif
            var elementList = grouped
                .Select(g =>
                {
                    if (level >= g.First().AncestorElements.Length)
                        return (object)(g.First().ContentAtom);
                    var ancestorBeingConstructed = g.First().AncestorElements[level];

                    if (ancestorBeingConstructed.Name == W.p)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentAtom.Name.ToString());
                        var newChildElements = groupedChildren
                            .Where(gc => gc.First().ContentAtom.Name != W.pPr)
                            .Select(gc =>
                            {
                                return CoalesceRecurse(gc, level + 1);
                            });
                        var newParaProps = groupedChildren
                            .Where(gc => gc.First().ContentAtom.Name == W.pPr)
                            .Select(gc => gc.Select(gce => gce.ContentAtom));
                        return new XElement(W.p,
                            ancestorBeingConstructed.Attributes(),
                            newParaProps, newChildElements);
                    }

                    if (ancestorBeingConstructed.Name == W.r)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentAtom.Name.ToString());
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                if (gc.First().ContentAtom.Name == W.t)
                                {
                                    var textOfTextElement = gc.Select(gce => gce.ContentAtom.Value).StringConcatenate();
                                    return (object)(new XElement(W.t, textOfTextElement));
                                }
                                else
                                    return gc.Select(gce => gce.ContentAtom);
                            });
                        var runProps = ancestorBeingConstructed.Elements(W.rPr);
                        return new XElement(W.r, runProps, newChildElements);
                    }

                    if (ancestorBeingConstructed.Name == W.tbl)
                        return ReconstructElement(g, ancestorBeingConstructed, W.tblPr, W.tblGrid);
                    if (ancestorBeingConstructed.Name == W.tr)
                        return ReconstructElement(g, ancestorBeingConstructed, W.trPr, null);
                    if (ancestorBeingConstructed.Name == W.tc)
                        return ReconstructElement(g, ancestorBeingConstructed, W.tcPr, null);
                    if (ancestorBeingConstructed.Name == W.sdt)
                        return ReconstructElement(g, ancestorBeingConstructed, W.sdtPr, W.sdtEndPr);
                    if (ancestorBeingConstructed.Name == W.sdtContent)
                        return ReconstructElement(g, ancestorBeingConstructed, null, null);

                    var newElement = new XElement(ancestorBeingConstructed.Name,
                        ancestorBeingConstructed.Attributes(),
                        CoalesceRecurse(g, level + 1));
                    return newElement;
                });
            return elementList;
        }

        private static XElement ReconstructElement(IGrouping<string, SplitRun> g, XElement ancestorBeingConstructed, XName props1XName,
            XName props2XName)
        {
            var groupedChildren = g
                .GroupAdjacent(gc => gc.ContentAtom.Name.ToString());
            var newChildElements = groupedChildren
                .Select(gc => gc.Select(gce => gce.ContentAtom));
            object props1 = null;
            if (props1XName != null)
                props1 = ancestorBeingConstructed.Elements(props1XName);
            object props2 = null;
            if (props2XName != null)
                props2 = ancestorBeingConstructed.Elements(props2XName);

            return new XElement(W.r, props1, props2, newChildElements);
        }

        private static void SplitAllElements(WordprocessingDocument wDoc)
        {
            foreach (var part in wDoc.ContentParts())
            {
                AnnotateAllElements(part);
                AnnotateWithSplitRuns(part);
            }
        }

        private static void AnnotateWithSplitRuns(OpenXmlPart part)
        {
            var partXDoc = part.GetXDocument();
            var splitRunsAnnotation = new SplitRunsAnnotation();
            XElement root = null;
            if (part is MainDocumentPart)
                root = partXDoc.Root.Element(W.body);
            else
                root = partXDoc.Root;

            var splitRuns = new List<SplitRun>();
            AnnotateWithSplitRunsRecurse(root, splitRuns);
            splitRunsAnnotation.SplitRuns = splitRuns.ToArray();
            
            part.AddAnnotation(splitRunsAnnotation);
        }

        private static void AnnotateWithSplitRunsRecurse(XElement element, List<SplitRun> splitRuns)
        {
            if (element.Name == W.body)
            {
                foreach (var item in element.Elements())
                    AnnotateWithSplitRunsRecurse(item, splitRuns);
                return;
            }

            if (element.Name == W.p)
            {
                var paraChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.pPr);
                foreach (var item in paraChildrenToProcess)
		            AnnotateWithSplitRunsRecurse(item, splitRuns);
                var paraProps = element.Element(W.pPr);
                if (paraProps == null)
                {
                    SplitRun pPrSplitRun = new SplitRun();
                    pPrSplitRun.ContentAtom = new XElement(W.pPr);
                    pPrSplitRun.AncestorElements = element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
                    splitRuns.Add(pPrSplitRun);
                }
                else
                {
                    SplitRun pPrSplitRun = new SplitRun();
                    pPrSplitRun.ContentAtom = paraProps;
                    pPrSplitRun.AncestorElements = element.Ancestors().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
                    splitRuns.Add(pPrSplitRun);
                }
                return;
            }

            if (element.Name == W.r)
            {
                var runChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.rPr);
                foreach (var item in runChildrenToProcess)
                    AnnotateWithSplitRunsRecurse(item, splitRuns);
                return;
            }

            if (element.Name == W.t)
            {
                var val = element.Value;
                foreach (var ch in val)
                {
                    var sr = new SplitRun();
                    sr.ContentAtom = new XElement(W.t, ch);
                    sr.AncestorElements = element.Ancestors().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
                    splitRuns.Add(sr);
                }
                return;
            }

            if (element.Name == W.tbl)
            {
                AnnotateElementWithProps(element, splitRuns, W.tblPr, W.tblGrid);
                return;
            }

            if (element.Name == W.tr)
            {
                AnnotateElementWithProps(element, splitRuns, W.trPr, null);
                return;
            }

            if (element.Name == W.tc)
            {
                AnnotateElementWithProps(element, splitRuns, W.tcPr, null);
                return;
            }

            if (element.Name == W.sdt)
            {
                AnnotateElementWithProps(element, splitRuns, W.sdtPr, null);
                return;
            }

            if (element.Name == W.sdtContent)
            {
                AnnotateElementWithProps(element, splitRuns, null, null);
                return;
            }

            if (element.Name == W.sectPr)
            {
                SplitRun sr3 = new SplitRun();
                sr3.ContentAtom = element;
                sr3.AncestorElements = element.Ancestors().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
                splitRuns.Add(sr3);
                return;
            }

            if (element.Name == W.proofErr ||
                element.Name == W.tblPr)
                return;

            SplitRun sr2 = new SplitRun();
            sr2.ContentAtom = element;
            sr2.AncestorElements = element.Ancestors().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
            splitRuns.Add(sr2);

            var elementChildrenToProcess = element
                .Elements();
            foreach (var item in elementChildrenToProcess)
                AnnotateWithSplitRunsRecurse(item, splitRuns);
        }

        private static void AnnotateElementWithProps(XElement element, List<SplitRun> splitRuns, XName props1XName, XName props2XName)
        {
            var runChildrenToProcess = element
                .Elements()
                .Where(e => e.Name != props1XName &&
                            e.Name != props2XName);
            foreach (var item in runChildrenToProcess)
                AnnotateWithSplitRunsRecurse(item, splitRuns);
        }

        private static void AnnotateAllElements(OpenXmlPart part)
        {
            var partXDoc = part.GetXDocument();
            int seq = 1;
            var content = partXDoc
                .Descendants()
                .Where(d =>
                    d.Name == W.p ||
                    d.Name == W.r ||
                    d.Name == W.tbl ||
                    d.Name == W.tr ||
                    d.Name == W.tc ||
                    d.Name == W.fldSimple ||
                    d.Name == W.hyperlink ||
                    d.Name == W.sdt ||
                    d.Name == W.smartTag);
            foreach (var d in content)
            {
                var unidKey = string.Format("{0:0000000000}", seq++);
                var newAtt = new XAttribute(PtOpenXml.Unid, unidKey);
                d.Add(newAtt);
            }
            var root = partXDoc.Root;
            if (root.Attribute(XNamespace.Xmlns + "pt14") == null)
            {
                root.Add(new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.pt.NamespaceName));
            }
            var ignorable = (string)root.Attribute(MC.Ignorable);
            if (ignorable != null)
            {
                var list = ignorable.Split(' ');
                if (!list.Contains("pt14"))
                {
                    ignorable += " pt14";
                    root.Attribute(MC.Ignorable).Value = ignorable;
                }
            }
            else
            {
                root.Add(new XAttribute(MC.Ignorable, "pt14"));
            }
            part.PutXDocument();
        }
    }

    internal class SplitRun
    {
        public XElement[] AncestorElements;
        public XElement ContentAtom;

        public string ToString(int indent)
        {
            int xNamePad = 16;
            var indentString = "".PadRight(indent);

            var sb = new StringBuilder();
            sb.Append(indentString);
            if (ContentAtom.Name == W.t)
            {
                sb.AppendFormat("{0}: {1} ", PadLocalName(xNamePad, this), ContentAtom.Value);
                AppendAncestorsDump(sb, this);
            }
            else
            {
                sb.AppendFormat("{0}:   ", PadLocalName(xNamePad, this));
                AppendAncestorsDump(sb, this);
            }
            return sb.ToString();
        }

        private static string PadLocalName(int xNamePad, SplitRun item)
        {
            return (item.ContentAtom.Name.LocalName + " ").PadRight(xNamePad, '-') + " ";
        }

        private void AppendAncestorsDump(StringBuilder sb, SplitRun sr)
        {
            var s = sr.AncestorElements.Select(p => p.Name.LocalName + GetUnid(p) + "/").StringConcatenate().TrimEnd('/');
            sb.Append("Ancestors:" + s);
        }

        private string GetUnid(XElement p)
        {
            var unid = (string)p.Attribute(PtOpenXml.Unid);
            if (unid == null)
                return "";
            return "[" + unid + "]";
        }
    }

    internal class SplitRunsAnnotation
    {
        public SplitRun[] SplitRuns;

        public string DumpSplitRunsAnnotation(int indent)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var item in SplitRuns)
                sb.Append(item.ToString(indent) + Environment.NewLine);
            return sb.ToString();
        }

    }
}
