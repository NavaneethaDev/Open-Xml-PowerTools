// TODO rationalize relationship ids in deleted content - need to copy the part over, update the relationship id in the deleted content.
// TODO make sure image modifications are captured

// prohibit
// - altChunk
// - permEnd
// - permStart
// - sdt
// - subDoc
// - smartTag
// - contentPart
//
// remove
// - proofErr
// - sectPr
//
// Test
// - fldSimple
// - hyperlink
// - endNotes
// - footNotes


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
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;
using System.Security.Cryptography;

namespace OpenXmlPowerTools
{
    public class WmlComparerSettings
    {
        public char[] WordSeparators;
        public string AuthorForRevisions = "Open-Xml-PowerTools";
        public string DateTimeForRevisions = DateTime.Now.ToString("o");

        public WmlComparerSettings()
        {
            // note that , and . are processed explicitly to handle cases where they are in a number or word
            WordSeparators = new[] { ' ', '-' }; // todo need to fix this for complete list
        }
    }

    public static class WmlComparer
    {
        public static bool s_DumpLog = false;

        // todo look for invalid content, throw if found
        public static WmlDocument Compare(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings)
        {
            WmlDocument wmlResult = new WmlDocument(source2);
            using (MemoryStream ms1 = new MemoryStream())
            using (MemoryStream ms2 = new MemoryStream())
            {
                ms1.Write(source1.DocumentByteArray, 0, source1.DocumentByteArray.Length);
                ms2.Write(source2.DocumentByteArray, 0, source2.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc1 = WordprocessingDocument.Open(ms1, true))
                using (WordprocessingDocument wDoc2 = WordprocessingDocument.Open(ms2, true))
                {
                    RemoveIrrelevantMarkup(wDoc1);
                    RemoveIrrelevantMarkup(wDoc2);

                    SimplifyMarkupSettings msSettings = new SimplifyMarkupSettings()
                    {
                        RemoveBookmarks = true,
                        AcceptRevisions = true,
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveFieldCodes = true,
                        RemoveGoBackBookmark = true,
                        RemoveLastRenderedPageBreak = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                    };
                    MarkupSimplifier.SimplifyMarkup(wDoc1, msSettings);
                    MarkupSimplifier.SimplifyMarkup(wDoc2, msSettings);

                    AddSha1HashToParagraphs(wDoc1);
                    AddSha1HashToParagraphs(wDoc2);
                    WmlRunSplitter.Split(wDoc1, new[] { wDoc1.MainDocumentPart });
                    WmlRunSplitter.Split(wDoc2, new[] { wDoc2.MainDocumentPart });

                    // if we were to compare headers and footers, then would want to iterate through ContentParts
                    //WmlRunSplitter.Split(wDoc1, wDoc1.ContentParts());
                    //WmlRunSplitter.Split(wDoc2, wDoc2.ContentParts());

                    ContentAtomListAnnotation sra1 = wDoc1.MainDocumentPart.Annotation<ContentAtomListAnnotation>();
                    ContentAtomListAnnotation sra2 = wDoc2.MainDocumentPart.Annotation<ContentAtomListAnnotation>();
                    return ApplyChanges(sra1, sra2, wmlResult, settings);
                }
            }
        }

        private static void RemoveIrrelevantMarkup(WordprocessingDocument wDoc)
        {
            wDoc.MainDocumentPart
                .GetXDocument()
                .Root
                .Descendants()
                .Where(d => d.Name == W.lastRenderedPageBreak ||
                            d.Name == W.bookmarkStart ||
                            d.Name == W.bookmarkEnd)
                .Remove();
            wDoc.MainDocumentPart.PutXDocument();
        }

        private static void AddSha1HashToParagraphs(WordprocessingDocument wDoc1)
        {
            var paragraphsToAnnotate = wDoc1.MainDocumentPart
                .GetXDocument()
                .Root
                .Descendants(W.p);

            foreach (var para in paragraphsToAnnotate)
            {
                var cloneParaForHashing = (XElement)CloneParaForHashing(wDoc1.MainDocumentPart, para);
                var s = cloneParaForHashing.ToString(SaveOptions.DisableFormatting)
                    .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                var ancestorsString = para.Ancestors().TakeWhile(a => a.Name != W.body).Select(a => a.Name.LocalName + "/").StringConcatenate();
                var sha1Hash = SHA1HashStringForUTF8String(s + ancestorsString);
                var pPr = para.Element(W.pPr);
                if (pPr == null)
                {
                    pPr = new XElement(W.pPr);
                    para.Add(pPr);
                }
                pPr.Add(new XAttribute(PtOpenXml.SHA1Hash, sha1Hash));
            }
        }

        /// <summary>
        /// Compute hash for string encoded as UTF8
        /// </summary>
        /// <param name="s">String to be hashed</param>
        /// <returns>40-character hex string</returns>
        private static string SHA1HashStringForUTF8String(string s)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(s);

            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);

            return HexStringFromBytes(hashBytes);
        }

        /// <summary>
        /// Convert an array of bytes to a string of hex digits
        /// </summary>
        /// <param name="bytes">array of bytes</param>
        /// <returns>String of hex digits</returns>
        private static string HexStringFromBytes(byte[] bytes)
        {
            var sb = new StringBuilder();
            foreach (byte b in bytes)
            {
                var hex = b.ToString("x2");
                sb.Append(hex);
            }
            return sb.ToString();
        }

        private static object CloneParaForHashing(OpenXmlPart mainDocumentPart, XNode node)
        {
            var element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.txbxContent ||
                    element.Name == W.bookmarkStart ||
                    element.Name == W.bookmarkEnd ||
                    element.Name == W.pPr ||
                    element.Name == W.rPr)
                    return null;

                if (element.Name == W.p)
                {
                    var newPara = new XElement(element.Name,
                        element.Attributes().Where(a => a.Name != W.rsid &&
                                a.Name != W.rsidDel &&
                                a.Name != W.rsidP &&
                                a.Name != W.rsidR &&
                                a.Name != W.rsidRDefault &&
                                a.Name != W.rsidRPr &&
                                a.Name != W.rsidSect &&
                                a.Name != W.rsidTr),
                        element.Nodes().Select(n => CloneParaForHashing(mainDocumentPart, n)));

                    var groupedRuns = newPara
                        .Elements()
                        .GroupAdjacent(e => e.Name == W.r &&
                            e.Elements().Count() == 1 &&
                            e.Element(W.t) != null);

                    var evenNewerPara = new XElement(element.Name,
                        groupedRuns.Select(g =>
                        {
                            if (g.Key)
                            {
                                var newRun = (object)new XElement(W.r,
                                    new XElement(W.t,
                                        g.Select(t => t.Value).StringConcatenate()));
                                return newRun;
                            }
                            return g;
                        }));

                    return evenNewerPara;
                }

                if (element.Name == W.pPr)
                {
                    var new_pPr = new XElement(W.pPr,
                        element.Attributes(),
                        element.Elements()
                            .Where(e => e.Name != W.sectPr)
                            .Select(n => CloneParaForHashing(mainDocumentPart, n)));
                    return new_pPr;
                }

                if (element.Name == W.r)
                {
                    var newRuns = element
                        .Elements()
                        .Where(e => e.Name != W.rPr)
                        .Select(rc => new XElement(W.r, CloneParaForHashing(mainDocumentPart, rc)));
                    return newRuns;
                }

                if (element.Name == A.blip)
                {
                    var newBlip = new XElement(element.Name,
                        element.Attributes().Select(a =>
                        {
                            if (!ComparisonUnit.s_RelationshipAttributeNames.Contains(a.Name))
                                return a;
                            var rId = (string)a;
                            OpenXmlPart oxp = mainDocumentPart.GetPartById(rId);
                            if (oxp == null)
                                throw new FileFormatException("Invalid WordprocessingML Document");
                            byte[] buffer = new byte[1024];
                            using (var str = oxp.GetStream())
                            {
                                var ret = str.Read(buffer, 0, buffer.Length);
                                if (ret == 0)
                                    throw new FileFormatException("Image contains no data");
                            }
                            var b64string = Convert.ToBase64String(buffer);
                            return new XAttribute(a.Name, b64string);
                        }),
                        element.Nodes().Select(n => CloneParaForHashing(mainDocumentPart, n)));
                    return newBlip;
                }

                if (element.Name == VML.shape)
                {
                    return new XElement(element.Name,
                        element.Attributes().Where(a => a.Name != "style"),
                        element.Nodes().Select(n => CloneParaForHashing(mainDocumentPart, n)));
                }

                if (element.Name == O.OLEObject)
                {
                    return new XElement(element.Name,
                        element.Attributes().Where(a =>
                            a.Name != "ObjectID" &&
                            a.Name != R.id),
                        element.Nodes().Select(n => CloneParaForHashing(mainDocumentPart, n)));
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CloneParaForHashing(mainDocumentPart, n)));
            }
            return node;
        }

        private static WmlDocument ApplyChanges(ContentAtomListAnnotation sra1, ContentAtomListAnnotation sra2, WmlDocument wmlResult,
            WmlComparerSettings settings)
        {
            var cu1 = GetComparisonUnitList(sra1, settings);
            var cu2 = GetComparisonUnitList(sra2, settings);

            if (s_DumpLog)
            {
                var sb3 = new StringBuilder();
                sb3.Append("ComparisonUnitList 1 =====" + Environment.NewLine);
                foreach (var item in cu1)
                {
                    sb3.Append("  ComparisonUnit =====" + Environment.NewLine);
                    foreach (var cu in item.Contents)
                    {
                        sb3.Append(cu.ToString(4) + Environment.NewLine);
                    }
                }
                sb3.Append("ComparisonUnitList 2 =====" + Environment.NewLine);
                foreach (var item in cu2)
                {
                    sb3.Append("  ComparisonUnit =====" + Environment.NewLine);
                    foreach (var cu in item.Contents)
                    {
                        sb3.Append(cu.ToString(4) + Environment.NewLine);
                    }
                }
                var sbs3 = sb3.ToString();
                Console.WriteLine(sbs3);
            }

            var correlatedSequence = Lcs(cu1, cu2);

            foreach (var cs in correlatedSequence.Where(z => z.CorrelationStatus == CorrelationStatus.Equal))
            {
                var zippedComparisonUnitArrays = cs.ComparisonUnitArray1.Zip(cs.ComparisonUnitArray2, (cuBefore, cuAfter) => new
                {
                    CuBefore = cuBefore,
                    CuAfter = cuAfter,
                });
                foreach (var cu in zippedComparisonUnitArrays)
                {
                    var zippedContents = cu.CuBefore.Contents.Zip(cu.CuAfter.Contents, (conBefore, conAfter) => new
                    {
                        ConBefore = conBefore,
                        ConAfter = conAfter,
                    });
                    foreach (var con in zippedContents)
                    {
                        var zippedAncestors = con.ConBefore.AncestorElements.Zip(con.ConAfter.AncestorElements, (ancBefore, ancAfter) => new
                        {
                            AncestorBefore = ancBefore,
                            AncestorAfter = ancAfter,
                        });
                        foreach (var anc in zippedAncestors)
                        {
                            if (anc.AncestorBefore == null || anc.AncestorAfter == null)
                                continue;
                            if (anc.AncestorBefore.Attribute(PtOpenXml.Unid) == null ||
                                anc.AncestorAfter.Attribute(PtOpenXml.Unid) == null)
                                continue;
                            var unid = (string)anc.AncestorAfter.Attribute(PtOpenXml.Unid);
                            if (anc.AncestorBefore.Attribute(PtOpenXml.Unid).Value != unid)
                                anc.AncestorBefore.Attribute(PtOpenXml.Unid).Value = unid;
                        }
                    }
                }
            }

            if (true)
            {
                var sb = new StringBuilder();
                foreach (var item in correlatedSequence)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }

            // the following gets a flattened list of ContentAtoms, with status indicated in each ContentAtom: Deleted, Inserted, or Equal
            var listOfContentAtoms = correlatedSequence
                .Select(cs =>
                {
                    if (cs.CorrelationStatus == CorrelationStatus.Equal)
                    {
                        var contentAtomList = cs
                            .ComparisonUnitArray2
                            .Select(cu => cu.Contents)
                            .SelectMany(m => m)
                            .Select(ca => new ContentAtom()
                            {
                                ContentElement = ca.ContentElement,
                                AncestorElements = ca.AncestorElements,
                                CorrelationStatus = CorrelationStatus.Equal,
                                Part = ca.Part,
                            });
                        return contentAtomList;
                    }
                    if (cs.CorrelationStatus == CorrelationStatus.Deleted)
                    {
                        var contentAtomList = cs
                            .ComparisonUnitArray1
                            .Select(cu => cu.Contents)
                            .SelectMany(m => m)
                            .Select(ca => new ContentAtom()
                            {
                                ContentElement = ca.ContentElement,
                                AncestorElements = ca.AncestorElements,
                                CorrelationStatus = CorrelationStatus.Deleted,
                                Part = ca.Part,
                            });
                        return contentAtomList;
                    }
                    else if (cs.CorrelationStatus == CorrelationStatus.Inserted)
                    {
                        var contentAtomList = cs
                            .ComparisonUnitArray2
                            .Select(cu => cu.Contents)
                            .SelectMany(m => m)
                            .Select(ca => new ContentAtom()
                            {
                                ContentElement = ca.ContentElement,
                                AncestorElements = ca.AncestorElements,
                                CorrelationStatus = CorrelationStatus.Inserted,
                                Part = ca.Part,
                            });
                        return contentAtomList;
                    }
                    else
                    {
                        throw new OpenXmlPowerToolsException("Internal error - should have no unknown correlated sequences at this point.");
                    }
                })
                .SelectMany(m => m)
                .ToList();

            if (true)
            {
                var sb2 = new StringBuilder();
                foreach (var item in listOfContentAtoms)
                    sb2.Append(item.ToString()).Append(Environment.NewLine);
                var sbs2 = sb2.ToString();
                Console.WriteLine(sbs2);
            }

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlResult.DocumentByteArray, 0, wmlResult.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var xDoc = wDoc.MainDocumentPart.GetXDocument();
                    var rootNamespaceAttributes = xDoc
                        .Root
                        .Attributes()
                        .Where(a => a.IsNamespaceDeclaration || a.Name.Namespace == MC.mc)
                        .ToList();

                    // ======================================
                    // The following produces a new valid WordprocessingML document from the listOfContentAtoms
                    XDocument newXDoc1 = ProduceNewXDocFromCorrelatedSequence(wDoc.MainDocumentPart, listOfContentAtoms, rootNamespaceAttributes, settings);

                    // little bit of cleanup
                    MoveLastSectPrToChildOfBody(newXDoc1);
                    XElement newXDoc2Root = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(newXDoc1.Root);
                    xDoc.Root.ReplaceWith(newXDoc2Root);

                    var root = xDoc.Root;
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
                    wDoc.MainDocumentPart.PutXDocument();
                }
                var updatedWmlResult = new WmlDocument("Dummy.docx", ms.ToArray());
                return updatedWmlResult;
            }
        }

#if false
        // leaving this code here, bc will need variation on this code when counting revisions.
        // not exactly, but close.  When counting revisions, need to coalesce adjacent revisions, with
        // probably certain exceptions like boundaries of tables.

        private static object CoalesceRunsInInsAndDel(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Elements().Any(c => c.Name == W.ins || c.Name == W.del))
                {
                    var groupedAdjacent = element
                        .Elements()
                        .GroupAdjacent(k =>
                        {
                            if (k.Name == W.ins || k.Name == W.del)
                                return k.Name.LocalName;
                            return "x";
                        })
                        .ToList();

                    var childElements = groupedAdjacent
                        .Select(g =>
                        {
                            if (g.Key == "x")
                                return (object)g;
                            // g.Key == "ins" || g.Key == "del"
                            var insOrDelGrouped = g
                                .GroupAdjacent(gc =>
                                {
                                    string key = "x";
                                    if (gc.Elements().Count() == 1 && gc.Elements(W.r).Count() == 1)
                                    {
                                        var firstElementName = gc.Elements().First().Name;
                                        key = firstElementName.LocalName + "|";
                                        var rPr = gc.Elements().First().Element(W.rPr);
                                        string rPrString = "";
                                        if (rPr != null)
                                            rPrString = rPr.ToString(SaveOptions.DisableFormatting);
                                        key += rPrString;
                                    }
                                    return key;
                                })
                                .ToList();
                            var newChildElements = insOrDelGrouped
                                .Select(idg =>
                                {
                                    if (idg.Key == "x")
                                        return (object)idg;
                                    XElement newChildElement = null;
                                    if (g.Key.StartsWith("ins"))
                                        newChildElement = new XElement(W.ins,
                                            g.First().Attributes());
                                    else
                                        newChildElement = new XElement(W.del,
                                            g.First().Attributes());
                                    var rPr = idg.Elements().Elements(W.rPr).FirstOrDefault();
                                    var run = new XElement(W.r,
                                        rPr,
                                        g.Elements().Elements().Where(e => e.Name != W.rPr));
                                    newChildElement.Add(run);
                                    return newChildElement;
                                })
                                .ToList();
                            return newChildElements;
                        })
                        .ToList();

                    var newElement = new XElement(element.Name,
                        new XAttribute(XNamespace.Xmlns + "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                        element.Attributes(),
                        childElements);
                    return newElement;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CoalesceRunsInInsAndDel(n)));
            }
            return node;
        }
#endif

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

        private static int s_MaxId = 0;

        private static XDocument ProduceNewXDocFromCorrelatedSequence(OpenXmlPart part, IEnumerable<ContentAtom> contentAtomList, List<XAttribute> rootNamespaceDeclarations, WmlComparerSettings settings)
        {
            // fabricate new MainDocumentPart from correlatedSequence

            if (s_DumpLog)
            {
                //dump out content atoms
                var sb = new StringBuilder();
                foreach (var item in contentAtomList)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }

            s_MaxId = 0;
            XDocument newXDoc = new XDocument();
            var newBodyChildren = CoalesceRecurse(part, contentAtomList, 0, settings);
            newXDoc.Add(
                new XElement(W.document,
                    rootNamespaceDeclarations,
                    new XElement(W.body, newBodyChildren)));

            var root = newXDoc.Root;
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

            return newXDoc;
        }

        private static object CoalesceRecurse(OpenXmlPart part, IEnumerable<ContentAtom> list, int level, WmlComparerSettings settings)
        {
            var grouped = list
                .GroupBy(ca =>
                {
                    // per the algorithm, I don't think that the following condition will ever evaluate to true
                    // for a table, we initially get all ContentAtoms for the entire table, then process.  When processing a row,
                    // no ContentAtoms will have ancestors outside the row.  Ditto for cells, and on down the tree.
                    if (level >= ca.AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 2 - why do we have ContentAtom objects with fewer ancestors than its siblings?");

                    // previously, instead of throwing, it returned a Guid to foce into their own group.
                    //return Guid.NewGuid().ToString().Replace("-", "");

                    var unid = (string)ca.AncestorElements[level].Attribute(PtOpenXml.Unid);
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
            }

            var elementList = grouped
                .Select(g =>
                {
                    // see the comment above at the beginning of CoalesceRecurse
                    if (level >= g.First().AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 1 - why do we have ContentAtom objects with fewer ancestors than its siblings?");

                    // previously, instead of throwing, it would return the content atom
                    // return (object)(g.First().ContentElement);

                    var ancestorBeingConstructed = g.First().AncestorElements[level];

                    if (ancestorBeingConstructed.Name == W.p)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentElement.Name.ToString() + " | " + gc.CorrelationStatus.ToString());
                        var newChildElements = groupedChildren
                            .Where(gc => gc.First().ContentElement.Name != W.pPr)
                            .Select(gc =>
                            {
                                if (gc.First().ContentElement.Name == M.oMath ||
                                    gc.First().ContentElement.Name == M.oMathPara)
                                {
                                    var deleting = g.First().CorrelationStatus == CorrelationStatus.Deleted;
                                    var inserting = g.First().CorrelationStatus == CorrelationStatus.Inserted;

                                    if (deleting)
                                    {
                                        return new XElement(W.del,
                                            new XAttribute(W.author, settings.AuthorForRevisions),
                                            new XAttribute(W.id, s_MaxId++),
                                            new XAttribute(W.date, settings.DateTimeForRevisions),
                                            gc.Select(gcc => gcc.ContentElement));
                                    }
                                    else if (inserting)
                                    {
                                        return new XElement(W.ins,
                                            new XAttribute(W.author, settings.AuthorForRevisions),
                                            new XAttribute(W.id, s_MaxId++),
                                            new XAttribute(W.date, settings.DateTimeForRevisions),
                                            gc.Select(gcc => gcc.ContentElement));
                                    }
                                    else
                                    {
                                        return gc.Select(gcc => gcc.ContentElement);
                                    }
                                }
                                return CoalesceRecurse(part, gc, level + 1, settings);
                            });

                        XElement pPr = null;
                        ContentAtom pPrContentAtom = null;
                        var newParaPropsGroup = groupedChildren
                            .FirstOrDefault(gc => gc.First().ContentElement.Name == W.pPr);
                        if (newParaPropsGroup != null)
                        {
                            pPrContentAtom = newParaPropsGroup.FirstOrDefault();
                            if (pPrContentAtom != null)
                            {
                                pPr = new XElement(pPrContentAtom.ContentElement); // clone so we can change it
                                if (pPrContentAtom.CorrelationStatus == CorrelationStatus.Deleted)
                                    pPr.Elements(W.sectPr).Remove(); // for now, don't move sectPr from old document to new document.
                            }
                        }
                        if (pPrContentAtom != null)
                        {
                            if (pPr == null)
                                pPr = new XElement(W.pPr);
                            if (pPrContentAtom.CorrelationStatus == CorrelationStatus.Deleted)
                            {
                                XElement rPr = pPr.Element(W.rPr);
                                if (rPr == null)
                                    rPr = new XElement(W.rPr);
                                rPr.Add(new XElement(W.del,
                                    new XAttribute(W.author, settings.AuthorForRevisions),
                                    new XAttribute(W.id, s_MaxId++),
                                    new XAttribute(W.date, settings.DateTimeForRevisions)));
                                if (pPr.Element(W.rPr) != null)
                                    pPr.Element(W.rPr).ReplaceWith(rPr);
                                else
                                    pPr.AddFirst(rPr);
                            }
                            else if (pPrContentAtom.CorrelationStatus == CorrelationStatus.Inserted)
                            {
                                XElement rPr = pPr.Element(W.rPr);
                                if (rPr == null)
                                    rPr = new XElement(W.rPr);
                                rPr.Add(new XElement(W.ins,
                                    new XAttribute(W.author, settings.AuthorForRevisions),
                                    new XAttribute(W.id, s_MaxId++),
                                    new XAttribute(W.date, settings.DateTimeForRevisions)));
                                if (pPr.Element(W.rPr) != null)
                                    pPr.Element(W.rPr).ReplaceWith(rPr);
                                else
                                    pPr.AddFirst(rPr);
                            }
                        }
                        var newPara = new XElement(W.p,
                            ancestorBeingConstructed.Attributes(),
                            pPr, newChildElements);
                        return newPara;
                    }

                    if (ancestorBeingConstructed.Name == W.r)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentElement.Name.ToString() + " | " + gc.CorrelationStatus.ToString());
                        var newChildElements = groupedChildren
                            .Select(gc =>
                            {
                                if (gc.First().ContentElement.Name == W.t)
                                {
                                    var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
                                    var del = gc.First().CorrelationStatus == CorrelationStatus.Deleted;
                                    var ins = gc.First().CorrelationStatus == CorrelationStatus.Inserted;
                                    if (del)
                                        return (object)(new XElement(W.delText,
                                            GetXmlSpaceAttribute(textOfTextElement),
                                            textOfTextElement));
                                    else
                                        return (object)(new XElement(W.t,
                                            GetXmlSpaceAttribute(textOfTextElement),
                                            textOfTextElement));
                                }
                                else
                                {
                                    var openXmlPartOfDeletedContent = gc.First().Part;
                                    var openXmlPartInNewDocument = part;
                                    return gc.Select(gce =>
                                        {
                                            //var del = gce.CorrelationStatus == CorrelationStatus.Deleted;
                                            //var ins = gce.CorrelationStatus == CorrelationStatus.Inserted;

                                            Package packageOfDeletedContent = openXmlPartOfDeletedContent.OpenXmlPackage.Package;
                                            Package packageOfNewContent = openXmlPartInNewDocument.OpenXmlPackage.Package;
                                            PackagePart partInDeletedDocument = packageOfDeletedContent.GetPart(part.Uri);
                                            PackagePart partInNewDocument = packageOfNewContent.GetPart(part.Uri);
                                            return MoveDeletedPartsToDestination(partInDeletedDocument, partInNewDocument, gce.ContentElement);
                                        });
                                }
                            });
                        var runProps = ancestorBeingConstructed.Elements(W.rPr);

                        var deleting = g.First().CorrelationStatus == CorrelationStatus.Deleted;
                        var inserting = g.First().CorrelationStatus == CorrelationStatus.Inserted;

                        if (deleting)
                        {
                            return new XElement(W.del,
                                new XAttribute(W.author, settings.AuthorForRevisions),
                                new XAttribute(W.id, s_MaxId++),
                                new XAttribute(W.date, settings.DateTimeForRevisions),
                                new XElement(W.r,
                                    runProps,
                                    newChildElements));
                        }
                        else if (inserting)
                        {
                            return new XElement(W.ins,
                                new XAttribute(W.author, settings.AuthorForRevisions),
                                new XAttribute(W.id, s_MaxId++),
                                new XAttribute(W.date, settings.DateTimeForRevisions),
                                new XElement(W.r,
                                    runProps,
                                    newChildElements));
                        }
                        else
                        {
                            return new XElement(W.r, runProps, newChildElements);
                        }
                    }

                    if (ancestorBeingConstructed.Name == W.tbl)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.tblPr, W.tblGrid, level, settings);
                    if (ancestorBeingConstructed.Name == W.tr)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.trPr, null, level, settings);
                    if (ancestorBeingConstructed.Name == W.tc)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.tcPr, null, level, settings);
                    if (ancestorBeingConstructed.Name == W.sdt)
                        return ReconstructElement(part, g, ancestorBeingConstructed, W.sdtPr, W.sdtEndPr, level, settings);
                    if (ancestorBeingConstructed.Name == W.sdtContent)
                        return (object)ReconstructElement(part, g, ancestorBeingConstructed, null, null, level, settings);

                    throw new OpenXmlPowerToolsException("Internal error - unrecognized ancestor being constructed.");
                    // previously, did the following, but should not be required.
                    //var newElement = new XElement(ancestorBeingConstructed.Name,
                    //    ancestorBeingConstructed.Attributes(),
                    //    CoalesceRecurse(g, level + 1));
                    //return newElement;
                })
                .ToList();
            return elementList;
        }

        private static object MoveDeletedPartsToDestination(PackagePart partOfDeletedContent, PackagePart partInNewDocument,
            XElement contentElement)
        {
            var elementsToUpdate = contentElement
                .Descendants()
                .Where(d => d.Attributes().Any(a => ComparisonUnit.s_RelationshipAttributeNames.Contains(a.Name)))
                .ToList();
            foreach (var element in elementsToUpdate)
            {
                var attributesToUpdate = element
                    .Attributes()
                    .Where(a => ComparisonUnit.s_RelationshipAttributeNames.Contains(a.Name))
                    .ToList();
                foreach (var att in attributesToUpdate)
                {
                    var rId = (string)att;


                    var relationshipForDeletedPart = partOfDeletedContent.GetRelationship(rId);
                    if (relationshipForDeletedPart == null)
                        throw new FileFormatException("Invalid document");

                    Uri targetUri = PackUriHelper
                        .ResolvePartUri(
                           new Uri(partOfDeletedContent.Uri.ToString(), UriKind.Relative),
                                 relationshipForDeletedPart.TargetUri);

                    var relatedPackagePart = partOfDeletedContent.Package.GetPart(targetUri);
                    var uriSplit = relatedPackagePart.Uri.ToString().Split('/');
                    var last = uriSplit[uriSplit.Length - 1].Split('.');
                    string uriString = null;
                    if (last.Length == 2)
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                            "P" + Guid.NewGuid().ToString().Replace("-", "") + "." + last[1];
                    }
                    else
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                            "P" + Guid.NewGuid().ToString().Replace("-", "");
                    }
                    Uri uri = null;
                    if (relatedPackagePart.Uri.IsAbsoluteUri)
                        uri = new Uri(uriString, UriKind.Absolute);
                    else
                        uri = new Uri(uriString, UriKind.Relative);

                    var newPart = partInNewDocument.Package.CreatePart(uri, relatedPackagePart.ContentType);
                    using (var oldPartStream = relatedPackagePart.GetStream())
                    using (var newPartStream = newPart.GetStream())
                        FileUtils.CopyStream(oldPartStream, newPartStream);

                    var newRid = "R" + Guid.NewGuid().ToString().Replace("-", "");
                    partInNewDocument.CreateRelationship(newPart.Uri, TargetMode.Internal, relationshipForDeletedPart.RelationshipType, newRid);
                    att.Value = newRid;

                    if (newPart.ContentType.EndsWith("xml"))
                    {
                        XDocument newPartXDoc = null;
                        using (var stream = newPart.GetStream())
                        {
                            newPartXDoc = XDocument.Load(stream);
                            MoveDeletedPartsToDestination(relatedPackagePart, newPart, newPartXDoc.Root);
                        }
                        using (var stream = newPart.GetStream())
                            newPartXDoc.Save(stream);
                    }
                }
            }
            return contentElement;
        }
#if false
        private static object MoveDeletedPartToDestination(OpenXmlPart openXmlPartOfDeletedContent, OpenXmlPart openXmlPartInNewDocument,
            XElement contentElement)
        {
            Package packageOfDeletedContent = openXmlPartOfDeletedContent.OpenXmlPackage.Package;
            Package packageOfNewContent = openXmlPartInNewDocument.OpenXmlPackage.Package;
            PackagePart partInNewDocument = packageOfNewContent.GetPart(openXmlPartInNewDocument.Uri);

            var elementsToUpdate = contentElement
                .Descendants()
                .Where(d => d.Attributes().Any(a => ComparisonUnit.s_RelationshipAttributeNames.Contains(a.Name)))
                .ToList();
            foreach (var element in elementsToUpdate)
            {
                var attributesToUpdate = element
                    .Attributes()
                    .Where(a => ComparisonUnit.s_RelationshipAttributeNames.Contains(a.Name))
                    .ToList();
                foreach (var att in attributesToUpdate)
                {
                    var rId = (string)att;


                    var relatedOpenXmlPart = openXmlPartOfDeletedContent.GetPartById(rId);
                    if (relatedOpenXmlPart == null)
                        throw new FileFormatException("Invalid document");
                    var relatedPackagePart = packageOfDeletedContent.GetPart(relatedOpenXmlPart.Uri);
                    var uriSplit = relatedOpenXmlPart.Uri.ToString().Split('/');
                    var last = uriSplit[uriSplit.Length - 1].Split('.');
                    string uriString = null;
                    if (last.Length == 2)
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                            "P" + Guid.NewGuid().ToString().Replace("-", "") + "." + last[1];
                    }
                    else
                    {
                        uriString = uriSplit.SkipLast(1).Select(p => p + "/").StringConcatenate() +
                            "P" + Guid.NewGuid().ToString().Replace("-", "");
                    }
                    Uri uri = null;
                    if (relatedOpenXmlPart.Uri.IsAbsoluteUri)
                        uri = new Uri(uriString, UriKind.Absolute);
                    else
                        uri = new Uri(uriString, UriKind.Relative);

                    var newPart = packageOfNewContent.CreatePart(uri, relatedPackagePart.ContentType); // not correct, need to make URI unique
                    using (var oldPartStream = relatedPackagePart.GetStream())
                    using (var newPartStream = newPart.GetStream())
                        FileUtils.CopyStream(oldPartStream, newPartStream);

                    var newRid = "R" + Guid.NewGuid().ToString().Replace("-", "");
                    partInNewDocument.CreateRelationship(newPart.Uri, TargetMode.Internal, relatedOpenXmlPart.RelationshipType, newRid);
                    att.Value = newRid;
                }
            }
            return contentElement;
        }
#endif

        private static XAttribute GetXmlSpaceAttribute(string textOfTextElement)
        {
            if (char.IsWhiteSpace(textOfTextElement[0]) ||
                char.IsWhiteSpace(textOfTextElement[textOfTextElement.Length - 1]))
                return new XAttribute(XNamespace.Xml + "space", "preserve");
            return null;
        }

        private static XElement ReconstructElement(OpenXmlPart part, IGrouping<string, ContentAtom> g, XElement ancestorBeingConstructed, XName props1XName,
            XName props2XName, int level, WmlComparerSettings settings)
        {
            var newChildElements = CoalesceRecurse(part, g, level + 1, settings);
            object props1 = null;
            if (props1XName != null)
                props1 = ancestorBeingConstructed.Elements(props1XName);
            object props2 = null;
            if (props2XName != null)
                props2 = ancestorBeingConstructed.Elements(props2XName);

            var reconstructedElement = new XElement(ancestorBeingConstructed.Name, props1, props2, newChildElements);
            return reconstructedElement;
        }

        private static List<CorrelatedSequence> Lcs(ComparisonUnit[] cu1, ComparisonUnit[] cu2)
        {
            // set up initial state - one CorrelatedSequence, UnKnown, contents == entire sequences (both)
            CorrelatedSequence cs = new CorrelatedSequence()
            {
                CorrelationStatus = OpenXmlPowerTools.CorrelationStatus.Unknown,
                ComparisonUnitArray1 = cu1,
                ComparisonUnitArray2 = cu2,
            };
            List<CorrelatedSequence> csList = new List<CorrelatedSequence>()
            {
                cs
            };

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var item in csList)
                    sb.Append(item.ToString());
                var s = sb.ToString();
                Console.WriteLine(s);
            }
            while (true)
            {
                var unknown = csList
                    .FirstOrDefault(z => z.CorrelationStatus == CorrelationStatus.Unknown);
                if (unknown == null)
                    break;

                // do LCS on paragraphs here
                List<CorrelatedSequence> newSequence = FindLongestCommonSequenceOfParagraphs(unknown);
                if (newSequence == null)
                    newSequence = FindLongestCommonSequence(unknown);

                var indexOfUnknown = csList.IndexOf(unknown);
                csList.Remove(unknown);

                newSequence.Reverse();
                foreach (var item in newSequence)
                    csList.Insert(indexOfUnknown, item);
            }

            return csList;
        }

        class ParagraphUnit
        {
            public List<ComparisonUnit> ComparisonUnits = new List<ComparisonUnit>();
            public string SHA1Hash = null;

            public override string ToString()
            {
                var sb = new StringBuilder();
                sb.Append("ParagraphUnit - SHA1Hash:" + SHA1Hash + Environment.NewLine);
                sb.Append(ComparisonUnit.ComparisonUnitListToString(this.ComparisonUnits.ToArray()));
                return sb.ToString();
            }
        }

        private static List<CorrelatedSequence> FindLongestCommonSequenceOfParagraphs(CorrelatedSequence unknown)
        {
            ParagraphUnit[] comparisonUnitArray1ByParagraphs = GetComparisonUnitListByParagraph(unknown.ComparisonUnitArray1);
            ParagraphUnit[] comparisonUnitArray2ByParagraphs = GetComparisonUnitListByParagraph(unknown.ComparisonUnitArray2);

            int lengthToCompare = Math.Min(comparisonUnitArray1ByParagraphs.Count(), comparisonUnitArray2ByParagraphs.Count());

            var countCommonParasAtBeginning = comparisonUnitArray1ByParagraphs
                .Take(lengthToCompare)
                .Zip(comparisonUnitArray2ByParagraphs, (pu1, pu2) =>
                {
                    return new
                    {
                        Pu1 = pu1,
                        Pu2 = pu2,
                    };
                })
                .TakeWhile(pair => pair.Pu1.SHA1Hash == pair.Pu2.SHA1Hash)
                .Count();

            var countCommonParasAtEnd = ((IEnumerable<ParagraphUnit>)comparisonUnitArray1ByParagraphs)
                .Skip(countCommonParasAtBeginning)
                .Reverse()
                .Take(lengthToCompare)
                .Zip(((IEnumerable<ParagraphUnit>)comparisonUnitArray2ByParagraphs).Reverse(), (pu1, pu2) =>
                {
                    return new
                    {
                        Pu1 = pu1,
                        Pu2 = pu2,
                    };
                })
                .TakeWhile(pair => pair.Pu1.SHA1Hash == pair.Pu2.SHA1Hash)
                .Count();

            List<CorrelatedSequence> newSequence = null;

            if (countCommonParasAtBeginning != 0 || countCommonParasAtEnd != 0)
            {
                newSequence = new List<CorrelatedSequence>();
                if (countCommonParasAtBeginning > 0)
                {
                    CorrelatedSequence cs = new CorrelatedSequence();
                    cs.CorrelationStatus = CorrelationStatus.Equal;
                    cs.ComparisonUnitArray1 = comparisonUnitArray1ByParagraphs
                        .Take(countCommonParasAtBeginning)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    cs.ComparisonUnitArray2 = comparisonUnitArray2ByParagraphs
                        .Take(countCommonParasAtBeginning)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    newSequence.Add(cs);
                }

                int middleSection1Len = comparisonUnitArray1ByParagraphs.Count() - countCommonParasAtBeginning - countCommonParasAtEnd;
                int middleSection2Len = comparisonUnitArray2ByParagraphs.Count() - countCommonParasAtBeginning - countCommonParasAtEnd;

                if (middleSection1Len > 0 && middleSection2Len == 0)
                {
                    CorrelatedSequence cs = new CorrelatedSequence();
                    cs.CorrelationStatus = CorrelationStatus.Deleted;
                    cs.ComparisonUnitArray1 = comparisonUnitArray1ByParagraphs
                        .Skip(countCommonParasAtBeginning)
                        .Take(middleSection1Len)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    cs.ComparisonUnitArray2 = null;
                    newSequence.Add(cs);
                }
                else if (middleSection1Len == 0 && middleSection2Len > 0)
                {
                    CorrelatedSequence cs = new CorrelatedSequence();
                    cs.CorrelationStatus = CorrelationStatus.Inserted;
                    cs.ComparisonUnitArray1 = null;
                    cs.ComparisonUnitArray2 = comparisonUnitArray2ByParagraphs
                        .Skip(countCommonParasAtBeginning)
                        .Take(middleSection2Len)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    newSequence.Add(cs);
                }
                else if (middleSection1Len > 0 && middleSection2Len > 0)
                {
                    CorrelatedSequence cs = new CorrelatedSequence();
                    cs.CorrelationStatus = CorrelationStatus.Unknown;
                    cs.ComparisonUnitArray1 = comparisonUnitArray1ByParagraphs
                        .Skip(countCommonParasAtBeginning)
                        .Take(middleSection1Len)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    cs.ComparisonUnitArray2 = comparisonUnitArray2ByParagraphs
                        .Skip(countCommonParasAtBeginning)
                        .Take(middleSection2Len)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    newSequence.Add(cs);
                }
                else if (middleSection1Len == 0 && middleSection2Len == 0)
                {
                    // nothing to do
                }

                if (countCommonParasAtEnd > 0)
                {
                    CorrelatedSequence cs = new CorrelatedSequence();
                    cs.CorrelationStatus = CorrelationStatus.Equal;
                    cs.ComparisonUnitArray1 = comparisonUnitArray1ByParagraphs
                        .Skip(countCommonParasAtBeginning)
                        .Skip(middleSection1Len)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    cs.ComparisonUnitArray2 = comparisonUnitArray2ByParagraphs
                        .Skip(countCommonParasAtBeginning)
                        .Skip(middleSection2Len)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    newSequence.Add(cs);
                }
            }
            else
            {
                var cul1 = comparisonUnitArray1ByParagraphs;
                var cul2 = comparisonUnitArray2ByParagraphs;
                int currentLongestCommonSequenceLength = 0;
                int currentI1 = -1;
                int currentI2 = -1;
                for (int i1 = 0; i1 < cul1.Length; i1++)
                {
                    for (int i2 = 0; i2 < cul2.Length; i2++)
                    {
                        var thisSequenceLength = 0;
                        var thisI1 = i1;
                        var thisI2 = i2;
                        while (true)
                        {
                            if (cul1[thisI1].SHA1Hash == cul2[thisI2].SHA1Hash)
                            {
                                thisI1++;
                                thisI2++;
                                thisSequenceLength++;
                                if (thisI1 == cul1.Length || thisI2 == cul2.Length)
                                {
                                    if (thisSequenceLength > currentLongestCommonSequenceLength)
                                    {
                                        currentLongestCommonSequenceLength = thisSequenceLength;
                                        currentI1 = i1;
                                        currentI2 = i2;
                                    }
                                    break;
                                }
                                continue;
                            }
                            else
                            {
                                if (thisSequenceLength > currentLongestCommonSequenceLength)
                                {
                                    currentLongestCommonSequenceLength = thisSequenceLength;
                                    currentI1 = i1;
                                    currentI2 = i2;
                                }
                                break;
                            }
                        }
                    }
                }

                var newListOfCorrelatedSequence = new List<CorrelatedSequence>();
                if (currentI1 == -1 && currentI2 == -1)
                {
                    return null;
                }

                if (currentI1 > 0 && currentI2 == 0)
                {
                    var deletedCorrelatedSequence = new CorrelatedSequence();
                    deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                    deletedCorrelatedSequence.ComparisonUnitArray1 = cul1.Take(currentI1).Select(cu => cu.ComparisonUnits).SelectMany(m => m).ToArray();
                    deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                    newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                }
                else if (currentI1 == 0 && currentI2 > 0)
                {
                    var insertedCorrelatedSequence = new CorrelatedSequence();
                    insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                    insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                    insertedCorrelatedSequence.ComparisonUnitArray2 = cul2
                        .Take(currentI2)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                }
                else if (currentI1 > 0 && currentI2 > 0)
                {
                    var unknownCorrelatedSequence = new CorrelatedSequence();
                    unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                    unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                        .Take(currentI1)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                        .Take(currentI2)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                }
                else if (currentI1 == 0 && currentI2 == 0)
                {
                    // nothing to do
                }

                var middleEqual = new CorrelatedSequence();
                middleEqual.CorrelationStatus = CorrelationStatus.Equal;
                middleEqual.ComparisonUnitArray1 = cul1
                    .Skip(currentI1)
                    .Take(currentLongestCommonSequenceLength)
                    .Select(cu => cu.ComparisonUnits)
                    .SelectMany(m => m)
                    .ToArray();
                middleEqual.ComparisonUnitArray2 = cul2
                    .Skip(currentI2)
                    .Take(currentLongestCommonSequenceLength)
                    .Select(cu => cu.ComparisonUnits)
                    .SelectMany(m => m)
                    .ToArray();
                newListOfCorrelatedSequence.Add(middleEqual);

                int endI1 = currentI1 + currentLongestCommonSequenceLength;
                int endI2 = currentI2 + currentLongestCommonSequenceLength;

                if (endI1 < cul1.Length && endI2 == cul2.Length)
                {
                    var deletedCorrelatedSequence = new CorrelatedSequence();
                    deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                    deletedCorrelatedSequence.ComparisonUnitArray1 = cul1
                        .Skip(endI1)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                    newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                }
                else if (endI1 == cul1.Length && endI2 < cul2.Length)
                {
                    var insertedCorrelatedSequence = new CorrelatedSequence();
                    insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                    insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                    insertedCorrelatedSequence.ComparisonUnitArray2 = cul2
                        .Skip(endI2)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                }
                else if (endI1 < cul1.Length && endI2 < cul2.Length)
                {
                    var unknownCorrelatedSequence = new CorrelatedSequence();
                    unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                    unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                        .Skip(endI1)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                        .Skip(endI2)
                        .Select(cu => cu.ComparisonUnits)
                        .SelectMany(m => m)
                        .ToArray();
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                }
                else if (endI1 == cul1.Length && endI2 == cul2.Length)
                {
                    // nothing to do
                }
                return newListOfCorrelatedSequence;
            }

            return newSequence;
        }

        private static ParagraphUnit[] GetComparisonUnitListByParagraph(ComparisonUnit[] comparisonUnit)
        {

            List<ParagraphUnit> listParaUnit = new List<ParagraphUnit>();
            ParagraphUnit thisParagraphUnit = new ParagraphUnit();
            foreach (var item in comparisonUnit)
            {
                if (item.Contents.First().ContentElement.Name == W.pPr)
                {
                    // note, the following RELIES on that the paragraph properties will only ever be in a group by themselves.
                    thisParagraphUnit.ComparisonUnits.Add(item);
                    thisParagraphUnit.SHA1Hash = (string)item.Contents.First().ContentElement.Attribute(PtOpenXml.SHA1Hash);
                    listParaUnit.Add(thisParagraphUnit);
                    thisParagraphUnit = new ParagraphUnit();
                    continue;
                }
                thisParagraphUnit.ComparisonUnits.Add(item);
            }
            if (thisParagraphUnit.ComparisonUnits.Any())
            {
                thisParagraphUnit.SHA1Hash = Guid.NewGuid().ToString();
                listParaUnit.Add(thisParagraphUnit);
            }
            return listParaUnit.ToArray();
        }

        private static List<CorrelatedSequence> FindLongestCommonSequence(CorrelatedSequence unknown)
        {
            var cul1 = unknown.ComparisonUnitArray1;
            var cul2 = unknown.ComparisonUnitArray2;
            int currentLongestCommonSequenceLength = 0;
            int currentI1 = -1;
            int currentI2 = -1;
            for (int i1 = 0; i1 < cul1.Length; i1++)
            {
                for (int i2 = 0; i2 < cul2.Length; i2++)
                {
                    var thisSequenceLength = 0;
                    var thisI1 = i1;
                    var thisI2 = i2;
                    while (true)
                    {
                        if (cul1[thisI1] == cul2[thisI2])
                        {
                            thisI1++;
                            thisI2++;
                            thisSequenceLength++;
                            if (thisI1 == cul1.Length || thisI2 == cul2.Length)
                            {
                                if (thisSequenceLength > currentLongestCommonSequenceLength)
                                {
                                    currentLongestCommonSequenceLength = thisSequenceLength;
                                    currentI1 = i1;
                                    currentI2 = i2;
                                }
                                break;
                            }
                            continue;
                        }
                        else
                        {
                            if (thisSequenceLength > currentLongestCommonSequenceLength)
                            {
                                currentLongestCommonSequenceLength = thisSequenceLength;
                                currentI1 = i1;
                                currentI2 = i2;
                            }
                            break;
                        }
                    }
                }
            }

            // if all we match is a paragraph mark, then don't match.
            if (currentLongestCommonSequenceLength == 1 && cul1[currentI1].Contents.First().ContentElement.Name == W.pPr)
            {
                currentLongestCommonSequenceLength = 0;
                currentI1 = -1;
                currentI2 = -1;
            }

            // if the paragraph mark is at the beginning of a LCS, then it is possible that erroneously matching a paragraph
            // mark that has been deleted.
            if (currentLongestCommonSequenceLength > 1 && cul1[currentI1].Contents.First().ContentElement.Name == W.pPr)
            {
                currentLongestCommonSequenceLength--;
                currentI1++;
                currentI2++;
            }

            // if the longest common subsequence starts with a space, and it is longer than 1, then don't include the space.
            if (currentI1 < cul1.Length && currentI1 != -1)
            {
                var contentElement = cul1[currentI1].Contents.First().ContentElement;
                if (currentLongestCommonSequenceLength > 1 && contentElement.Name == W.t && char.IsWhiteSpace(contentElement.Value[0]))
                {
                    currentI1++;
                    currentI2++;
                    currentLongestCommonSequenceLength--;
                }
            }

            // if the longest common subsequence is only a space, and it is only a single char long, then don't match
            if (currentLongestCommonSequenceLength == 1 && currentI1 < cul1.Length && currentI1 != -1)
            {
                var contentElement = cul1[currentI1].Contents.First().ContentElement;
                if (contentElement.Name == W.t && char.IsWhiteSpace(contentElement.Value[0]))
                {
                    currentLongestCommonSequenceLength = 0;
                    currentI1 = -1;
                    currentI2 = -1;
                }
            }

            // if the longest common subsequence length is less than 20% of the entire length, then don't match
            var max = Math.Max(cul1.Length, cul2.Length);
            if (((decimal)currentLongestCommonSequenceLength / (decimal)max) < 0.1M)
            {
                currentLongestCommonSequenceLength = 0;
                currentI1 = -1;
                currentI2 = -1;
            }

            var newListOfCorrelatedSequence = new List<CorrelatedSequence>();
            if (currentI1 == -1 && currentI2 == -1)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1;
                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);

                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2;
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);

                return newListOfCorrelatedSequence;
            }

            if (currentI1 > 0 && currentI2 == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1.Take(currentI1).ToArray();
                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);

                var equalCorrelatedSequence = new CorrelatedSequence();
                equalCorrelatedSequence.CorrelationStatus = CorrelationStatus.Equal;
                equalCorrelatedSequence.ComparisonUnitArray1 = cul1.Skip(currentI1).Take(currentLongestCommonSequenceLength).ToArray();
                equalCorrelatedSequence.ComparisonUnitArray2 = cul2.Take(currentLongestCommonSequenceLength).ToArray();
                newListOfCorrelatedSequence.Add(equalCorrelatedSequence);
            }
            else if (currentI1 == 0 && currentI2 > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2.Take(currentI2).ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);

                var equalCorrelatedSequence = new CorrelatedSequence();
                equalCorrelatedSequence.CorrelationStatus = CorrelationStatus.Equal;
                equalCorrelatedSequence.ComparisonUnitArray1 = cul1.Take(currentLongestCommonSequenceLength).ToArray();
                equalCorrelatedSequence.ComparisonUnitArray2 = cul2.Skip(currentI2).Take(currentLongestCommonSequenceLength).ToArray();
                newListOfCorrelatedSequence.Add(equalCorrelatedSequence);
            }
            else if (currentI1 > 0 && currentI2 > 0)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1.Take(currentI1).ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2.Take(currentI2).ToArray();
                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);

                var equalCorrelatedSequence = new CorrelatedSequence();
                equalCorrelatedSequence.CorrelationStatus = CorrelationStatus.Equal;
                equalCorrelatedSequence.ComparisonUnitArray1 = cul1.Skip(currentI1).Take(currentLongestCommonSequenceLength).ToArray();
                equalCorrelatedSequence.ComparisonUnitArray2 = cul2.Skip(currentI2).Take(currentLongestCommonSequenceLength).ToArray();
                newListOfCorrelatedSequence.Add(equalCorrelatedSequence);
            }
            else if (currentI1 == 0 && currentI2 == 0)
            {
                var equalCorrelatedSequence = new CorrelatedSequence();
                equalCorrelatedSequence.CorrelationStatus = CorrelationStatus.Equal;
                equalCorrelatedSequence.ComparisonUnitArray1 = cul1.Take(currentLongestCommonSequenceLength).ToArray();
                equalCorrelatedSequence.ComparisonUnitArray2 = cul2.Take(currentLongestCommonSequenceLength).ToArray();
                newListOfCorrelatedSequence.Add(equalCorrelatedSequence);
            }

            int endI1 = currentI1 + currentLongestCommonSequenceLength;
            int endI2 = currentI2 + currentLongestCommonSequenceLength;

            if (endI1 < cul1.Length && endI2 == cul2.Length)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1.Skip(endI1).ToArray();
                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
            }
            else if (endI1 == cul1.Length && endI2 < cul2.Length)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2.Skip(endI2).ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (endI1 < cul1.Length && endI2 < cul2.Length)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1.Skip(endI1).ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2.Skip(endI2).ToArray();
                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (endI1 == cul1.Length && endI2 == cul2.Length)
            {
                // nothing to do here...
            }
            return newListOfCorrelatedSequence;
        }

        private static XName[] WordBreakElements = new XName[] {
            W.pPr,
            W.tab,
            W.br,
            W.continuationSeparator,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.drawing,
            W.endnoteRef,
            W.footnoteRef,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W._object,
            W.ptab,
            W.separator,
            W.sym,
            W.yearLong,
            W.yearShort,
            M.oMathPara,
            M.oMath,
        };

        private static ComparisonUnit[] GetComparisonUnitList(ContentAtomListAnnotation contentAtomListAnnotation, WmlComparerSettings settings)
        {
            var contentAtomList = contentAtomListAnnotation.ContentAtomList;
            var groupingKey = contentAtomList
                .Select((sr, i) =>
                {
                    string key = null;
                    if (sr.ContentElement.Name == W.t)
                    {
                        string chr = sr.ContentElement.Value;
                        var ch = chr[0];
                        if (ch == '.' || ch == ',')
                        {
                            bool beforeIsDigit = false;
                            if (i > 0)
                            {
                                var prev = contentAtomList[i - 1];
                                if (prev.ContentElement.Name == W.t && char.IsDigit(prev.ContentElement.Value[0]))
                                    beforeIsDigit = true;
                            }
                            bool afterIsDigit = false;
                            if (i < contentAtomList.Length - 1)
                            {
                                var next = contentAtomList[i + 1];
                                if (next.ContentElement.Name == W.t && char.IsDigit(next.ContentElement.Value[0]))
                                    afterIsDigit = true;
                            }
                            if (beforeIsDigit || afterIsDigit)
                            {
                                key = "x | ";
                                var ancestorsKey = sr
                                    .AncestorElements
                                    .Where(a => a.Name != W.r)
                                    .Select(a => (string)a.Attribute(PtOpenXml.Unid) + "-")
                                    .StringConcatenate();
                                key += ancestorsKey;
                            }
                            else
                                key = i.ToString();
                        }
                        else if (settings.WordSeparators.Contains(ch))
                            key = i.ToString();
                        else
                        {
                            key = "x | ";
                            var ancestorsKey = sr
                                .AncestorElements
                                .Where(a => a.Name != W.r)
                                .Select(a => (string)a.Attribute(PtOpenXml.Unid) + "-")
                                .StringConcatenate();
                            key += ancestorsKey;
                        }
                    }
                    else if (WordBreakElements.Contains(sr.ContentElement.Name))
                    {
                        key = i.ToString();
                    }
                    else
                    {
                        key = "x | ";
                        var ancestorsKey = sr
                            .AncestorElements
                            .Where(a => a.Name != W.r)
                            .Select(a => (string)a.Attribute(PtOpenXml.Unid) + "-")
                            .StringConcatenate();
                        key += ancestorsKey;
                    }
                    return new
                    {
                        Key = key,
                        ContentAtomMember = sr
                    };
                });

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var item in groupingKey)
                {
                    sb.Append(item.Key + Environment.NewLine);
                    sb.Append("    " + item.ContentAtomMember.ToString(0) + Environment.NewLine);
                }
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }
            var groupedByWords = groupingKey
                .GroupAdjacent(gc => gc.Key);

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var group in groupedByWords)
                {
                    sb.Append("Group ===== " + group.Key + Environment.NewLine);
                    foreach (var gc in group)
                    {
                        sb.Append("    " + gc.ContentAtomMember.ToString(0) + Environment.NewLine);
                    }
                }
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }

            ComparisonUnit[] cul = groupedByWords
                .Select(g => new ComparisonUnit(g.Select(gc => gc.ContentAtomMember)))
                .ToArray();

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var group in cul)
                {
                    sb.Append("Group ===== " + Environment.NewLine);
                    foreach (var gc in group.Contents)
                    {
                        sb.Append("    " + gc.ToString(0) + Environment.NewLine);
                    }
                }
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }
            return cul;
        }
    }

    internal class ComparisonUnit : IEquatable<ComparisonUnit>
    {
        public List<ContentAtom> Contents;
        public ComparisonUnit(IEnumerable<ContentAtom> contentAtomList)
        {
            Contents = contentAtomList.ToList();
        }

        public static XName[] s_RelationshipAttributeNames = new XName[] {
            R.embed,
            R.link,
            R.id,
            R.cs,
            R.dm,
            R.lo,
            R.qs,
            R.href,
            R.pict,
        };

        public string ToString(int indent)
        {
            var sb = new StringBuilder();
            foreach (var contentAtom in Contents)
                sb.Append(contentAtom.ToString(indent) + Environment.NewLine);
            return sb.ToString();
        }

        public static string ComparisonUnitListToString(ComparisonUnit[] comparisonUnit)
        {
            var sb = new StringBuilder();
            sb.Append("Dumping ComparisonUnit List" + Environment.NewLine);
            for (int i = 0; i < comparisonUnit.Length; i++)
            {
                sb.AppendFormat("  Comparison Unit: {0}", i).Append(Environment.NewLine);
                foreach (var su in comparisonUnit[i].Contents)
                {
                    sb.Append(su.ToString(4));
                    sb.Append(Environment.NewLine);
                }
            }
            var sbs = sb.ToString();
            return sbs;
        }

        // TODO need to add all other elements that we should discard.
        // end notes?
        // foot notes?
        // go through standard, look for other things to ignore.
        private static XName[] s_ElementsToIgnoreWhenComparing = new[] {
            W.bookmarkStart,
            W.bookmarkEnd,
            W.commentRangeStart,
            W.commentRangeEnd,
            W.proofErr,
        };

        public bool Equals(ComparisonUnit other)
        {
            if (other == null)
                return false;

            if (this.Contents.Any(c => c.ContentElement.Name == W.t) ||
                other.Contents.Any(c => c.ContentElement.Name == W.t))
            {
                var txt1 = this
                    .Contents
                    .Where(c => c.ContentElement.Name == W.t)
                    .Select(c => c.ContentElement.Value)
                    .StringConcatenate();
                var txt2 = other
                    .Contents
                    .Where(c => c.ContentElement.Name == W.t)
                    .Select(c => c.ContentElement.Value)
                    .StringConcatenate();
                if (txt1 != txt2)
                    return false;

                var seq1 = this
                    .Contents
                    .Where(c => ! s_ElementsToIgnoreWhenComparing.Contains(c.ContentElement.Name));
                var seq2 = other
                    .Contents
                    .Where(c => ! s_ElementsToIgnoreWhenComparing.Contains(c.ContentElement.Name));
                if (seq1.Count() != seq2.Count())
                    return false;
                var zipped = seq1.Zip(seq2, (s1, s2) => new
                {
                    Cu1 = s1,
                    Cu2 = s2,
                });
                var anyNotEqual = (zipped.Any(z =>
                    {
                        var a1 = z.Cu1.AncestorElements.Select(a => a.Name.ToString() + "|").StringConcatenate();
                        var a2 = z.Cu2.AncestorElements.Select(a => a.Name.ToString() + "|").StringConcatenate();
                        return a1 != a2;
                    }));
                if (anyNotEqual)
                    return false;
                return true;
            }
            else
            {
                var seq1 = this
                    .Contents
                    .Where(c => !s_ElementsToIgnoreWhenComparing.Contains(c.ContentElement.Name));
                var seq2 = other
                    .Contents
                    .Where(c => !s_ElementsToIgnoreWhenComparing.Contains(c.ContentElement.Name));
                if (seq1.Count() != seq2.Count())
                    return false;
                
                var zipped = seq1.Zip(seq2, (s1, s2) => new
                {
                    Cu1 = s1,
                    Cu2 = s2,
                });
                var anyNotEqual = (zipped.Any(z =>
                {
                    if (z.Cu1.ContentElement.Name != z.Cu2.ContentElement.Name)
                        return true;
                    var a1 = z.Cu1.AncestorElements.Select(a => a.Name.ToString() + "|").StringConcatenate();
                    var a2 = z.Cu2.AncestorElements.Select(a => a.Name.ToString() + "|").StringConcatenate();
                    if (a1 != a2)
                        return true;
                    var name = z.Cu1.ContentElement.Name;
                    if (name == M.oMath || name == M.oMathPara)
                    {
                        var equ = XNode.DeepEquals(z.Cu1.ContentElement, z.Cu2.ContentElement);
                        return !equ;
                    }
                    if (name == W.drawing)
                    {
                        var relationshipIds1 = z.Cu1.ContentElement
                            .Descendants()
                            .Attributes()
                            .Where(a => s_RelationshipAttributeNames.Contains(a.Name))
                            .Select(a => (string)a)
                            .ToList();
                        var relationshipIds2 = z.Cu2.ContentElement
                            .Descendants()
                            .Attributes()
                            .Where(a => s_RelationshipAttributeNames.Contains(a.Name))
                            .Select(a => (string)a)
                            .ToList();
                        if (relationshipIds1.Count() != relationshipIds2.Count())
                            return true;
                        var sourcePart1 = this.Contents.First().Part;
                        var sourcePart2 = other.Contents.First().Part;
                        var zipped2 = relationshipIds1.Zip(relationshipIds2, (rid1, rid2) =>
                            {
                                return new
                                {
                                    RelId1 = rid1,
                                    RelId2 = rid2,
                                };
                            });
                        foreach (var pair in zipped2)
                        {
                            var oxp1 = sourcePart1.GetPartById(pair.RelId1);
                            if (oxp1 == null)
                                throw new FileFormatException("Invalid WordprocessingML Document");
                            var oxp2 = sourcePart2.GetPartById(pair.RelId2);
                            if (oxp2 == null)
                                throw new FileFormatException("Invalid WordprocessingML Document");
                            byte[] buffer1 = new byte[1024];
                            byte[] buffer2 = new byte[1024];
                            using (var str1 = oxp1.GetStream())
                            using (var str2 = oxp2.GetStream())
                            {
                                var ret1 = str1.Read(buffer1, 0, buffer1.Length);
                                var ret2 = str2.Read(buffer2, 0, buffer2.Length);
                                if (ret1 == 0 && ret2 == 0)
                                    continue;
                                if (ret1 != ret2)
                                    return true;
                                for (int i = 0; i < buffer1.Length; i++)
                                    if (buffer1[i] != buffer2[i])
                                        return true;
                                continue;
                            }
                        }
                        return false;
                    }
                    return false;
                }));
                if (anyNotEqual)
                    return false;
                return true;
            }
        }

        public override bool Equals(Object obj)
        {
            if (obj == null)
                return false;

            ComparisonUnit cuObj = obj as ComparisonUnit;
            if (cuObj == null)
                return false;
            else
                return Equals(cuObj);
        }

        public override int GetHashCode()
        {
            return this.GetHashCode();
        }

        public static bool operator ==(ComparisonUnit comparisonUnit1, ComparisonUnit comparisonUnit2)
        {
            if (((object)comparisonUnit1) == null || ((object)comparisonUnit2) == null)
                return Object.Equals(comparisonUnit1, comparisonUnit2);

            return comparisonUnit1.Equals(comparisonUnit2);
        }

        public static bool operator !=(ComparisonUnit comparisonUnit1, ComparisonUnit comparisonUnit2)
        {
            if (((object)comparisonUnit1) == null || ((object)comparisonUnit2) == null)
                return !Object.Equals(comparisonUnit1, comparisonUnit2);

            return !(comparisonUnit1.Equals(comparisonUnit2));
        }
    }

    enum CorrelationStatus
    {
        Nil,
        Unknown,
        Inserted,
        Deleted,
        Equal,
    }

    class CorrelatedSequence
    {
        public CorrelationStatus CorrelationStatus;

        // if ComparisonUnitList1 == null and ComparisonUnitList2 contains sequence, then inserted content.
        // if ComparisonUnitList2 == null and ComparisonUnitList1 contains sequence, then deleted content.
        // if ComparisonUnitList2 contains sequence and ComparisonUnitList1 contains sequence, then either is Unknown or Equal.
        public ComparisonUnit[] ComparisonUnitArray1;
        public ComparisonUnit[] ComparisonUnitArray2;

        public override string ToString()
        {
            var sb = new StringBuilder();
            var indentString = "  ";
            var indentString4 = "    ";
            sb.Append("CorrelatedSequence =====" + Environment.NewLine);
            sb.Append(indentString + "CorrelatedItem =====" + Environment.NewLine);
            sb.Append(indentString4 + "CorrelationStatus: " + CorrelationStatus.ToString() + Environment.NewLine);
            if (ComparisonUnitArray1 != null)
            {
                sb.Append(indentString4 + "ComparisonUnitList1 =====" + Environment.NewLine);
                foreach (var item in ComparisonUnitArray1)
                    sb.Append(item.ToString(6) + Environment.NewLine);
            }
            if (ComparisonUnitArray2 != null)
            {
                sb.Append(indentString + "ComparisonUnitList2 =====" + Environment.NewLine);
                foreach (var item in ComparisonUnitArray2)
                    sb.Append(item.ToString(6) + Environment.NewLine);
            }
            return sb.ToString();
        }
    }
}
