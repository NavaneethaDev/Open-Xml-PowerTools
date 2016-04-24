﻿#define SHORT_UNID

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
    internal class WmlContentAtomList
    {
        public static bool s_DumpLog = false;

        public static void CreateContentAtomList(WordprocessingDocument wDoc, OpenXmlPart part)
        {
            AssignIdToAllElements(part);  // add the Guid id to every element for which we need to establish identity
            MoveLastSectPrIntoLastParagraph(part);
            AnnotatePartWithContentAtomListAnnotation(part); // adds the list of ContentAtom objects as an annotation to the part
        }

        internal static XDocument Coalesce(ContentAtomListAnnotation contentAtomListAnnotation)
        {
            XDocument newXDoc = new XDocument();
            var newBodyChildren = CoalesceRecurse(contentAtomListAnnotation.ContentAtomList, 0);
            newXDoc.Add(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.pt.NamespaceName),
                new XElement(W.body, newBodyChildren)));

            // little bit of cleanup
            MoveLastSectPrToChildOfBody(newXDoc);
            XElement newXDoc2Root = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(newXDoc.Root);
            newXDoc.Root.ReplaceWith(newXDoc2Root);
            return newXDoc;
        }

        private static void MoveLastSectPrToChildOfBody(XDocument newXDoc)
        {
            var lastParaWithSectPr = newXDoc
                .Root
                .Elements(W.body)
                .Elements(W.p)
                .Where(p => p.Elements(W.pPr).Elements(W.sectPr).Any())
                .LastOrDefault();
            if (lastParaWithSectPr != null)
            {
                newXDoc.Root.Element(W.body).Add(lastParaWithSectPr.Elements(W.pPr).Elements(W.sectPr));
                lastParaWithSectPr.Elements(W.pPr).Elements(W.sectPr).Remove();
            }
        }

        private static object CoalesceRecurse(IEnumerable<ContentAtom> list, int level)
        {
            var grouped = list
                .GroupBy(sr =>
                {
                    // per the algorithm, The following condition will never evaluate to true
                    // if it evaluates to true, then the basic mechanism for breaking a hierarchical structure into flat and back is broken.

                    // for a table, we initially get all ContentAtoms for the entire table, then process.  When processing a row,
                    // no ContentAtoms will have ancestors outside the row.  Ditto for cells, and on down the tree.
                    if (level >= sr.AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 4 - why do we have ContentAtom objects with fewer ancestors than its siblings?");

                    var unid = (string)sr.AncestorElements[level].Attribute(PtOpenXml.Unid);
                    return unid;
                });

            if (s_DumpLog)
            {
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
                var sbs = sb.ToString();
            }
            
            var elementList = grouped
                .Select(g =>
                {
                    // see the comment above at the beginning of CoalesceRecurse
                    if (level >= g.First().AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 3 - why do we have ContentAtom objects with fewer ancestors than its siblings?");

                    var ancestorBeingConstructed = g.First().AncestorElements[level];

                    if (ancestorBeingConstructed.Name == W.p)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentElement.Name.ToString());
                        var newChildElements = groupedChildren
                            .Where(gc => gc.First().ContentElement.Name != W.pPr)
                            .Select(gc =>
                            {
                                return CoalesceRecurse(gc, level + 1);
                            });
                        var newParaProps = groupedChildren
                            .Where(gc => gc.First().ContentElement.Name == W.pPr)
                            .Select(gc => gc.Select(gce => gce.ContentElement));
                        return new XElement(W.p,
                            ancestorBeingConstructed.Attributes(),
                            newParaProps, newChildElements);
                    }

                    if (ancestorBeingConstructed.Name == W.r)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentElement.Name.ToString());
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                if (gc.First().ContentElement.Name == W.t)
                                {
                                    var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
                                    return (object)(new XElement(W.t,
                                        GetXmlSpaceAttribute(textOfTextElement),
                                        textOfTextElement));
                                }
                                else
                                    return gc.Select(gce => gce.ContentElement);
                            });
                        var runProps = ancestorBeingConstructed.Elements(W.rPr);
                        return new XElement(W.r, runProps, newChildElements);
                    }

                    if (ancestorBeingConstructed.Name == W.tbl)
                        return ReconstructElement(g, ancestorBeingConstructed, W.tblPr, W.tblGrid, level);
                    if (ancestorBeingConstructed.Name == W.tr)
                        return ReconstructElement(g, ancestorBeingConstructed, W.trPr, null, level);
                    if (ancestorBeingConstructed.Name == W.tc)
                        return ReconstructElement(g, ancestorBeingConstructed, W.tcPr, null, level);
                    if (ancestorBeingConstructed.Name == W.sdt)
                        return ReconstructElement(g, ancestorBeingConstructed, W.sdtPr, W.sdtEndPr, level);
                    if (ancestorBeingConstructed.Name == W.sdtContent)
                        return ReconstructElement(g, ancestorBeingConstructed, null, null, level);

                    var newElement = new XElement(ancestorBeingConstructed.Name,
                        ancestorBeingConstructed.Attributes(),
                        CoalesceRecurse(g, level + 1));
                    return newElement;
                })
                .ToList();
            return elementList;
        }

        private static XAttribute GetXmlSpaceAttribute(string textOfTextElement)
        {
            if (char.IsWhiteSpace(textOfTextElement[0]) ||
                char.IsWhiteSpace(textOfTextElement[textOfTextElement.Length - 1]))
                return new XAttribute(XNamespace.Xml + "space", "preserve");
            return null;
        }

        private static XElement ReconstructElement(IGrouping<string, ContentAtom> g, XElement ancestorBeingConstructed, XName props1XName,
            XName props2XName, int level)
        {
            var newChildElements = CoalesceRecurse(g, level + 1);
            object props1 = null;
            if (props1XName != null)
                props1 = ancestorBeingConstructed.Elements(props1XName);
            object props2 = null;
            if (props2XName != null)
                props2 = ancestorBeingConstructed.Elements(props2XName);

            var reconstructedElement = new XElement(ancestorBeingConstructed.Name, props1, props2, newChildElements);
            return reconstructedElement;
        }

        private static void MoveLastSectPrIntoLastParagraph(OpenXmlPart part)
        {
            XDocument xDoc = part.GetXDocument();
            var lastSectPrList = xDoc.Root.Element(W.body).Elements(W.sectPr).ToList();
            if (lastSectPrList.Count() > 1)
                throw new OpenXmlPowerToolsException("Invalid document");
            var lastSectPr = lastSectPrList.FirstOrDefault();
            if (lastSectPr != null)
            {
                var lastParagraph = xDoc.Root.Elements(W.body).Elements(W.p).LastOrDefault();
                if (lastParagraph == null)
                    throw new OpenXmlPowerToolsException("Invalid document");
                var pPr = lastParagraph.Element(W.pPr);
                if (pPr == null)
                {
                    pPr = new XElement(W.pPr);
                    lastParagraph.AddFirst(W.pPr);
                }
                pPr.Add(lastSectPr);
                xDoc.Root.Element(W.body).Elements(W.sectPr).Remove();
            }
        }

        private static void AnnotatePartWithContentAtomListAnnotation(OpenXmlPart part)
        {
            var partXDoc = part.GetXDocument();
            var contentAtomListAnnotation = new ContentAtomListAnnotation();
            XElement root = null;
            if (part is MainDocumentPart)
                root = partXDoc.Root.Element(W.body);
            else
                root = partXDoc.Root;

            var contentAtomList = new List<ContentAtom>();
            AnnotateWithContentAtomListRecurse(part, root, contentAtomList);
            contentAtomListAnnotation.ContentAtomList = contentAtomList.ToArray();

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var ca in contentAtomListAnnotation.ContentAtomList)
                {
                    sb.Append(ca.ToString(0)).Append(Environment.NewLine);
                }
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }

            part.AddAnnotation(contentAtomListAnnotation);
        }

        // note that if we were to support comments, this would change
        private static XName[] AllowableRunChildren = new XName[] {
            W.br,
            W.drawing,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.endnoteRef,
            W.footnoteRef,
            W.footnoteReference,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W._object,
            W.pgNum,
            W.ptab,
            W.softHyphen,
            W.sym,
            W.tab,
            W.yearLong,
            W.yearShort,
            M.oMathPara,
            M.oMath,
            W.fldChar,
            W.instrText,
        };

        private static XName[] ElementsToThrowAway = new XName[] {
            W.bookmarkStart,
            W.bookmarkEnd,
            W.lastRenderedPageBreak,
            W.proofErr,
            W.tblPr,
            W.sectPr,
        };

        private static void AnnotateWithContentAtomListRecurse(OpenXmlPart part, XElement element, List<ContentAtom> contentAtomList)
        {
            if (element.Name == W.body)
            {
                foreach (var item in element.Elements())
                    AnnotateWithContentAtomListRecurse(part, item, contentAtomList);
                return;
            }

            if (element.Name == W.p)
            {
                var paraChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.pPr);
                foreach (var item in paraChildrenToProcess)
		            AnnotateWithContentAtomListRecurse(part, item, contentAtomList);
                var paraProps = element.Element(W.pPr);
                if (paraProps == null)
                {
                    ContentAtom pPrContentAtom = new ContentAtom();
                    pPrContentAtom.ContentElement = new XElement(W.pPr);
                    pPrContentAtom.Part = part;
                    pPrContentAtom.AncestorElements = element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
                    contentAtomList.Add(pPrContentAtom);
                }
                else
                {
                    ContentAtom pPrContentAtom = new ContentAtom();
                    pPrContentAtom.ContentElement = paraProps;
                    pPrContentAtom.Part = part;
                    pPrContentAtom.AncestorElements = element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
                    contentAtomList.Add(pPrContentAtom);
                }
                return;
            }

            if (element.Name == W.r)
            {
                var runChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.rPr);
                foreach (var item in runChildrenToProcess)
                    AnnotateWithContentAtomListRecurse(part, item, contentAtomList);
                return;
            }

            if (element.Name == W.t)
            {
                var val = element.Value;
                foreach (var ch in val)
                {
                    var sr = new ContentAtom();
                    sr.ContentElement = new XElement(W.t, ch);
                    sr.Part = part;
                    sr.AncestorElements = element.Ancestors().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
                    contentAtomList.Add(sr);
                }
                return;
            }

            if (AllowableRunChildren.Contains(element.Name))
            {
                ContentAtom sr3 = new ContentAtom();
                sr3.ContentElement = element;
                sr3.Part = part;
                sr3.AncestorElements = element.Ancestors().TakeWhile(a => a.Name != W.body).Reverse().ToArray();
                contentAtomList.Add(sr3);
                return;
            }

            if (element.Name == W.tbl)
            {
                AnnotateElementWithProps(part, element, contentAtomList, W.tblPr, W.tblGrid, W.tblPrEx);
                return;
            }

            if (element.Name == W.tr)
            {
                AnnotateElementWithProps(part, element, contentAtomList, W.trPr, W.tblPrEx, null);
                return;
            }

            if (element.Name == W.tc)
            {
                AnnotateElementWithProps(part, element, contentAtomList, W.tcPr, W.tblPrEx, null);
                return;
            }

            if (element.Name == W.sdt)
            {
                AnnotateElementWithProps(part, element, contentAtomList, W.sdtPr, null, null);
                return;
            }

            if (element.Name == W.sdtContent)
            {
                AnnotateElementWithProps(part, element, contentAtomList, null, null, null);
                return;
            }

            if (element.Name == W.hyperlink)
            {
                AnnotateElementWithProps(part, element, contentAtomList, null, null, null);
                return;
            }

            if (ElementsToThrowAway.Contains(element.Name))
                return;

            throw new OpenXmlPowerToolsException("Internal error - unexpected element");
        }

        private static void AnnotateElementWithProps(OpenXmlPart part, XElement element, List<ContentAtom> contentAtomList, XName props1XName, XName props2XName, XName props3XName)
        {
            var runChildrenToProcess = element
                .Elements()
                .Where(e => e.Name != props1XName &&
                            e.Name != props2XName &&
                            e.Name != props3XName);
            foreach (var item in runChildrenToProcess)
                AnnotateWithContentAtomListRecurse(part, item, contentAtomList);
        }

        private static XName[] ElementsToHaveUnid = new XName[]
        {
            W.p,
            W.r,
            W.tbl,
            W.tr,
            W.tc,
            W.fldSimple,
            W.hyperlink,
            W.sdt,
            W.smartTag,
        };

        private static void AssignIdToAllElements(OpenXmlPart part)
        {
            var partXDoc = part.GetXDocument();
            var content = partXDoc
                .Descendants()
                .Where(d => ElementsToHaveUnid.Contains(d.Name));
            foreach (var d in content)
            {
                var newAtt = new XAttribute(PtOpenXml.Unid, Guid.NewGuid().ToString().Replace("-", "")
#if SHORT_UNID
                    .Substring(0, 12) // when debugging
#endif
                );
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

    internal class ContentAtom
    {
        public XElement[] AncestorElements;
        public XElement ContentElement;
        public OpenXmlPart Part;
        public CorrelationStatus CorrelationStatus;

        public string ToString(int indent)
        {
            int xNamePad = 16;
            var indentString = "".PadRight(indent);

            var sb = new StringBuilder();
            sb.Append(indentString);
            string correlationStatus = "";
            if (CorrelationStatus != OpenXmlPowerTools.CorrelationStatus.Nil)
                correlationStatus = string.Format("({0}) ", CorrelationStatus.ToString().PadRight(8));
            if (ContentElement.Name == W.t)
            {
                sb.AppendFormat("{0}: {1} {2}", PadLocalName(xNamePad, this), ContentElement.Value, correlationStatus);
                AppendAncestorsDump(sb, this);
            }
            else
            {
                sb.AppendFormat("{0}:   {1}", PadLocalName(xNamePad, this), correlationStatus);
                AppendAncestorsDump(sb, this);
            }
            return sb.ToString();
        }

        public override string ToString()
        {
            return ToString(0);
        }

        private static string PadLocalName(int xNamePad, ContentAtom item)
        {
            return (item.ContentElement.Name.LocalName + " ").PadRight(xNamePad, '-') + " ";
        }

        private void AppendAncestorsDump(StringBuilder sb, ContentAtom sr)
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

    internal class ContentAtomListAnnotation
    {
        public ContentAtom[] ContentAtomList;

        public string DumpContentAtomListAnnotation(int indent)
        {
            StringBuilder sb = new StringBuilder();
            var cal = ContentAtomList
                .Select((ca, i) => new
                {
                    ContentAtom = ca,
                    Index = i,
                });
            foreach (var item in cal)
                sb.Append("".PadRight(indent))
                  .AppendFormat("[{0:000000}] ", item.Index + 1)
                  .Append(item.ContentAtom.ToString(0) + Environment.NewLine);
            return sb.ToString();
        }

    }
}