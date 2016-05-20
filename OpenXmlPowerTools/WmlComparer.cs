//#define SHORT_UNID
#undef SHORT_UNID

// Test
// - endNotes
// - footNotes
// - ptab is not adequately tested.

/***************************************************************************

Copyright (c) Microsoft Corporation 2016.

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
        public double DetailThreshold = 0.15;

        public WmlComparerSettings()
        {
            // note that , and . are processed explicitly to handle cases where they are in a number or word
            WordSeparators = new[] { ' ', '-' }; // todo need to fix this for complete list
        }
    }

    public static class WmlComparer
    {
        public static bool s_DumpLog = false;
        public static bool s_True = true;

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
                    TestForInvalidContent(wDoc1);
                    TestForInvalidContent(wDoc2);
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
                        RemoveHyperlinks = true,
                    };
                    MarkupSimplifier.SimplifyMarkup(wDoc1, msSettings);
                    MarkupSimplifier.SimplifyMarkup(wDoc2, msSettings);

                    AddSha1HashToBlockLevelContent(wDoc1);
                    AddSha1HashToBlockLevelContent(wDoc2);
                    var cal1 = WmlComparer.CreateComparisonUnitAtomList(wDoc1, wDoc1.MainDocumentPart).ToArray();
                    var cus1 = GetComparisonUnitList(cal1, settings);
                    var cal2 = WmlComparer.CreateComparisonUnitAtomList(wDoc2, wDoc2.MainDocumentPart).ToArray();
                    var cus2 = GetComparisonUnitList(cal2, settings);

                    return ApplyChanges(cus1, cus2, wmlResult, settings);
                }
            }
        }

        // prohibit
        // - altChunk
        // - subDoc
        // - contentPart
        private static void TestForInvalidContent(WordprocessingDocument wDoc)
        {
            foreach (var part in wDoc.ContentParts())
            {
                var xDoc = part.GetXDocument();
                if (xDoc.Descendants(W.altChunk).Any())
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:altChunk");
                if (xDoc.Descendants(W.subDoc).Any())
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:subDoc");
                if (xDoc.Descendants(W.contentPart).Any())
                    throw new OpenXmlPowerToolsException("Unsupported document, contains w:contentPart");
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
            wDoc.MainDocumentPart
                .GetXDocument()
                .Root
                .Descendants()
                .Attributes()
                .Where(a => a.Name.Namespace == PtOpenXml.pt)
                .Remove();
            wDoc.MainDocumentPart.PutXDocument();
        }

        private static void AddSha1HashToBlockLevelContent(WordprocessingDocument wDoc)
        {
            var blockLevelContentToAnnotate = wDoc.MainDocumentPart
                .GetXDocument()
                .Root
                .Descendants()
                .Where(d => ElementsToHaveSha1Hash.Contains(d.Name));

            foreach (var blockLevelContent in blockLevelContentToAnnotate)
            {
                var cloneBlockLevelContentForHashing = (XElement)CloneBlockLevelContentForHashing(wDoc.MainDocumentPart, blockLevelContent);
                var shaString = cloneBlockLevelContentForHashing.ToString(SaveOptions.DisableFormatting)
                    .Replace(" xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
                var sha1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaString);
                blockLevelContent.Add(new XAttribute(PtOpenXml.SHA1Hash, sha1Hash));
            }
        }

        static XName[] AttributesToTrimWhenCloning = new XName[] {
            WP14.anchorId,
            WP14.editId,
        };

        private static object CloneBlockLevelContentForHashing(OpenXmlPart mainDocumentPart, XNode node)
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
                    var clonedPara = new XElement(element.Name,
                        element.Attributes().Where(a => a.Name != W.rsid &&
                                a.Name != W.rsidDel &&
                                a.Name != W.rsidP &&
                                a.Name != W.rsidR &&
                                a.Name != W.rsidRDefault &&
                                a.Name != W.rsidRPr &&
                                a.Name != W.rsidSect &&
                                a.Name != W.rsidTr),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));

                    var groupedRuns = clonedPara
                        .Elements()
                        .GroupAdjacent(e => e.Name == W.r &&
                            e.Elements().Count() == 1 &&
                            e.Element(W.t) != null);

                    var clonedParaWithGroupedRuns = new XElement(element.Name,
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

                    return clonedParaWithGroupedRuns;
                }

                if (element.Name == W.pPr)
                {
                    var cloned_pPr = new XElement(W.pPr,
                        element.Attributes(),
                        element.Elements()
                            .Where(e => e.Name != W.sectPr)
                            .Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));
                    return cloned_pPr;
                }

                if (element.Name == W.r)
                {
                    var clonedRuns = element
                        .Elements()
                        .Where(e => e.Name != W.rPr)
                        .Select(rc => new XElement(W.r, CloneBlockLevelContentForHashing(mainDocumentPart, rc)));
                    return clonedRuns;
                }

                if (element.Name == W.tbl)
                {
                    var clonedTable = new XElement(W.tbl,
                        element.Elements(W.tr).Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));
                    return clonedTable;
                }

                if (element.Name == W.tr)
                {
                    var clonedRow = new XElement(W.tr,
                        element.Elements(W.tc).Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));
                    return clonedRow;
                }

                if (element.Name == W.tc)
                {
                    var clonedCell = new XElement(W.tc,
                        element.Elements().Where(z => z.Name != W.tcPr).Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));
                    return clonedCell;
                }

                if (ComparisonUnitWord.s_ElementsWithRelationshipIds.Contains(element.Name))
                {
                    var newElement = new XElement(element.Name,
                        element.Attributes().Where(a => !AttributesToTrimWhenCloning.Contains(a.Name)).Select(a =>
                        {
                            if (!ComparisonUnitWord.s_RelationshipAttributeNames.Contains(a.Name))
                                return a;
                            var rId = (string)a;
                            OpenXmlPart oxp = mainDocumentPart.GetPartById(rId);
                            if (oxp == null)
                                throw new FileFormatException("Invalid WordprocessingML Document");

                            var anno = oxp.Annotation<PartSHA1HashAnnotation>();
                            if (anno != null)
                                return new XAttribute(a.Name, anno.Hash);

                            if (!oxp.ContentType.EndsWith("xml"))
                            {
                                using (var str = oxp.GetStream())
                                {
                                    byte[] ba;
                                    using (BinaryReader br = new BinaryReader(str))
                                    {
                                        ba = br.ReadBytes((int)str.Length);
                                    }
                                    var sha1 = WmlComparerUtil.SHA1HashStringForByteArray(ba);
                                    oxp.AddAnnotation(new PartSHA1HashAnnotation(sha1));
                                    return new XAttribute(a.Name, sha1);
                                }
                            }
                            return null;
                        }),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));
                    return newElement;
                }

                if (element.Name == VML.shape)
                {
                    return new XElement(element.Name,
                        element.Attributes().Where(a => a.Name != "style"),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));
                }

                if (element.Name == O.OLEObject)
                {
                    return new XElement(element.Name,
                        element.Attributes().Where(a =>
                            a.Name != "ObjectID" &&
                            a.Name != R.id),
                        element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));
                }

                return new XElement(element.Name,
                    element.Attributes().Where(a => !AttributesToTrimWhenCloning.Contains(a.Name)),
                    element.Nodes().Select(n => CloneBlockLevelContentForHashing(mainDocumentPart, n)));
            }
            return node;
        }


        private static List<CorrelatedSequence> FindCommonAtBeginningAndEnd(CorrelatedSequence unknown, WmlComparerSettings settings)
        {
            int lengthToCompare = Math.Min(unknown.ComparisonUnitArray1.Length, unknown.ComparisonUnitArray2.Length);

            var countCommonAtBeginning = unknown
                .ComparisonUnitArray1
                .Take(lengthToCompare)
                .Zip(unknown.ComparisonUnitArray2,
                    (pu1, pu2) =>
                    {
                        return new
                        {
                            Pu1 = pu1,
                            Pu2 = pu2,
                        };
                    })
                    .TakeWhile(pair => pair.Pu1.SHA1Hash == pair.Pu2.SHA1Hash)
                    .Count();

            if (countCommonAtBeginning != 0 && ((double)countCommonAtBeginning / (double)lengthToCompare) < settings.DetailThreshold)
                countCommonAtBeginning = 0;

            var countCommonAtEnd = unknown
                .ComparisonUnitArray1
                .Skip(countCommonAtBeginning)
                .Reverse()
                .Take(lengthToCompare)
                .Zip(unknown
                    .ComparisonUnitArray2
                    .Skip(countCommonAtBeginning)
                    .Reverse()
                    .Take(lengthToCompare),
                    (pu1, pu2) =>
                    {
                        return new
                        {
                            Pu1 = pu1,
                            Pu2 = pu2,
                        };
                    })
                    .TakeWhile(pair => pair.Pu1.SHA1Hash == pair.Pu2.SHA1Hash)
                    .Count();

            // never start a common section with a paragraph mark.  However, it is OK to set two paragraph marks as equal.
            while (true)
            {
                if (countCommonAtEnd <= 1)
                    break;

                var firstCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .LastOrDefault();

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                if (firstCommonWord == null)
                    break;

                // if the word contains more than one atom, then not a paragraph mark
                if (firstCommonWord.Contents.Count() != 1)
                    break;

                var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                if (firstCommonAtom == null)
                    break;

                if (firstCommonAtom.ContentElement.Name != W.pPr)
                    break;

                countCommonAtEnd--;
            }

            bool isOnlyParagraphMark = false;
            if (countCommonAtEnd == 1)
            {
                var firstCommon = unknown
                    .ComparisonUnitArray1
                    .Reverse()
                    .Take(countCommonAtEnd)
                    .LastOrDefault();

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                if (firstCommonWord != null)
                {
                    // if the word contains more than one atom, then not a paragraph mark
                    if (firstCommonWord.Contents.Count() == 1)
                    {
                        var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                        if (firstCommonAtom != null)
                        {
                            if (firstCommonAtom.ContentElement.Name == W.pPr)
                                isOnlyParagraphMark = true;
                        }
                    }
                }
            }

            if (!isOnlyParagraphMark && countCommonAtEnd != 0 && ((double)countCommonAtEnd / (double)lengthToCompare) < settings.DetailThreshold)
                countCommonAtEnd = 0;

            if (countCommonAtBeginning == 0 && countCommonAtEnd == 0)
                return null;

            var newSequence = new List<CorrelatedSequence>();
            if (countCommonAtBeginning != 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Equal;

                cs.ComparisonUnitArray1 = unknown
                    .ComparisonUnitArray1
                    .Take(countCommonAtBeginning)
                    .ToArray();

                cs.ComparisonUnitArray2 = unknown
                    .ComparisonUnitArray2
                    .Take(countCommonAtBeginning)
                    .ToArray();

                newSequence.Add(cs);
            }

            var middleLeft = unknown
                .ComparisonUnitArray1
                .Skip(countCommonAtBeginning)
                .SkipLast(countCommonAtEnd)
                .ToArray();

            var middleRight = unknown
                .ComparisonUnitArray2
                .Skip(countCommonAtBeginning)
                .SkipLast(countCommonAtEnd)
                .ToArray();

            if (middleLeft.Length > 0 && middleRight.Length == 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Deleted;
                cs.ComparisonUnitArray1 = middleLeft;
                cs.ComparisonUnitArray2 = null;
                newSequence.Add(cs);
            }
            else if (middleLeft.Length == 0 && middleRight.Length > 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Inserted;
                cs.ComparisonUnitArray1 = null;
                cs.ComparisonUnitArray2 = middleRight;
                newSequence.Add(cs);
            }
            else if (middleLeft.Length > 0 && middleRight.Length > 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Unknown;
                cs.ComparisonUnitArray1 = middleLeft;
                cs.ComparisonUnitArray2 = middleRight;
                newSequence.Add(cs);
            }

            if (countCommonAtEnd != 0)
            {
                CorrelatedSequence cs = new CorrelatedSequence();
                cs.CorrelationStatus = CorrelationStatus.Equal;

                cs.ComparisonUnitArray1 = unknown
                    .ComparisonUnitArray1
                    .Skip(countCommonAtBeginning + middleLeft.Length)
                    .ToArray();

                cs.ComparisonUnitArray2 = unknown
                    .ComparisonUnitArray2
                    .Skip(countCommonAtBeginning + middleRight.Length)
                    .ToArray();

                newSequence.Add(cs);
            }
            return newSequence;
        }

        private static WmlDocument ApplyChanges(ComparisonUnit[] cu1, ComparisonUnit[] cu2, WmlDocument wmlResult,
            WmlComparerSettings settings)
        {
            if (s_DumpLog)
            {
                var sb3 = new StringBuilder();
                sb3.Append("ComparisonUnitList 1 =====" + Environment.NewLine + Environment.NewLine);
                sb3.Append(ComparisonUnit.ComparisonUnitListToString(cu1));
                sb3.Append(Environment.NewLine);
                sb3.Append("ComparisonUnitList 2 =====" + Environment.NewLine + Environment.NewLine);
                sb3.Append(ComparisonUnit.ComparisonUnitListToString(cu2));
                var sbs3 = sb3.ToString();
                Console.WriteLine(sbs3);
            }

            var correlatedSequence = Lcs(cu1, cu2, settings);

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var item in correlatedSequence)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
                //TestUtil.NotePad(sbs);
            }

            // for any deleted or inserted rows, we go into the w:trPr properties, and add the appropriate w:ins or w:del element, and therefore
            // when generating the document, the appropriate row will be marked as deleted or inserted.

            foreach (var dcs in correlatedSequence.Where(cs =>
                cs.CorrelationStatus == CorrelationStatus.Deleted || cs.CorrelationStatus == CorrelationStatus.Inserted))
            {
                // iterate through all deleted/inserted items in dcs.ComparisonUnitArray1/ComparisonUnitArray2
                var toIterateThrough = dcs.ComparisonUnitArray1;
                if (dcs.CorrelationStatus == CorrelationStatus.Inserted)
                    toIterateThrough = dcs.ComparisonUnitArray2;

                foreach (var ca in toIterateThrough)
                {
                    var cug = ca as ComparisonUnitGroup;
                    
                    // this works because we will never see a table in this list, only rows.  If tables were in this list, would need to recursively
                    // go into children, but tables are always flattened in the LCS process.

                    // when we have a row, it is only necessary to find the first content atom of the row, then find the row ancestor, and then tweak
                    // the w:trPr

                    if (cug != null && cug.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        var firstContentAtom = cug.DescendantContentAtoms().FirstOrDefault();
                        if (firstContentAtom == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        var tr = firstContentAtom
                            .AncestorElements
                            .Reverse()
                            .FirstOrDefault(a => a.Name == W.tr);

                        if (tr == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        var trPr = tr.Element(W.trPr);
                        if (trPr == null)
                        {
                            trPr = new XElement(W.trPr);
                            tr.AddFirst(trPr);
                        }
                        XName revTrackElementName = null;
                        if (dcs.CorrelationStatus == CorrelationStatus.Deleted)
                            revTrackElementName = W.del;
                        else if (dcs.CorrelationStatus == CorrelationStatus.Inserted)
                            revTrackElementName = W.ins;
                        trPr.Add(new XElement(revTrackElementName,
                            new XAttribute(W.author, settings.AuthorForRevisions),
                            new XAttribute(W.id, s_MaxId++),
                            new XAttribute(W.date, settings.DateTimeForRevisions)));
                    }
                }
            }

            // the following gets a flattened list of ComparisonUnitAtoms, with status indicated in each ComparisonUnitAtom: Deleted, Inserted, or Equal

            var listOfComparisonUnitAtoms = correlatedSequence
                .Select(cs =>
                {
                    if (cs.CorrelationStatus == CorrelationStatus.Equal)
                    {
                        var comparisonUnitAtomList = cs
                            .ComparisonUnitArray2
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m)
                            .Select(ca =>
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part)
                                {
                                    CorrelationStatus = CorrelationStatus.Equal,
                                });
                        return comparisonUnitAtomList;
                    }
                    else if (cs.CorrelationStatus == CorrelationStatus.Deleted)
                    {
                        var comparisonUnitAtomList = cs
                            .ComparisonUnitArray1
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m)
                            .Select(ca =>
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part)
                                {
                                    CorrelationStatus = CorrelationStatus.Deleted,
                                });
                        return comparisonUnitAtomList;
                    }
                    else if (cs.CorrelationStatus == CorrelationStatus.Inserted)
                    {
                        var comparisonUnitAtomList = cs
                            .ComparisonUnitArray2
                            .Select(ca => ca.DescendantContentAtoms())
                            .SelectMany(m => m)
                            .Select(ca =>
                                new ComparisonUnitAtom(ca.ContentElement, ca.AncestorElements, ca.Part)
                                {
                                    CorrelationStatus = CorrelationStatus.Inserted,
                                });
                        return comparisonUnitAtomList;
                    }
                    else
                        throw new OpenXmlPowerToolsException("Internal error");
                })
                .SelectMany(m => m)
                .ToList();

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var item in listOfComparisonUnitAtoms)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
                //TestUtil.NotePad(sbs);
            }

            // hack = set the guid ID of the table, row, or cell from the 'before' document to be equal to the 'after' document.

            // note - we don't want to do the hack until after flattening all of the groups.  At the end of the flattening, we should simply
            // have a list of ComparisonUnitAtoms, appropriately marked as equal, inserted, or deleted.

            // the table id will be hacked in the normal course of events.
            // in the case where a row is deleted, not necessary to hack - the deleted row ID will do.
            // in the case where a row is inserted, not necessary to hack - the inserted row ID will do as well.

            HashSet<string> alreadySetUnids = new HashSet<string>();
            foreach (var cs in correlatedSequence.Where(z => z.CorrelationStatus == CorrelationStatus.Equal))
            {
                var zippedComparisonUnitArrays = cs.ComparisonUnitArray1.Zip(cs.ComparisonUnitArray2, (cuBefore, cuAfter) => new
                {
                    CuBefore = cuBefore,
                    CuAfter = cuAfter,
                });
                foreach (var cu in zippedComparisonUnitArrays)
                {
                    var beforeDescendantContentAtoms = cu.CuBefore
                        .DescendantContentAtoms();

                    var afterDescendantContentAtoms = cu.CuAfter
                        .DescendantContentAtoms();

                    var zippedContents = beforeDescendantContentAtoms
                        .Zip(afterDescendantContentAtoms,
                            (conBefore, conAfter) => new
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
                            var beforeUnid = (string)anc.AncestorBefore.Attribute(PtOpenXml.Unid);
                            var afterUnid = (string)anc.AncestorAfter.Attribute(PtOpenXml.Unid);
                            if (beforeUnid != afterUnid)
                            {
                                if (!alreadySetUnids.Contains(beforeUnid))
                                {
                                    alreadySetUnids.Add(beforeUnid);
                                    anc.AncestorBefore.Attribute(PtOpenXml.Unid).Value = afterUnid;
                                }
                            }
                        }
                    }
                }
            }

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var item in listOfComparisonUnitAtoms)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                // TestUtil.NotePad(sbs);
            }

            // and then finally can generate the document with revisions

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
                     // The following produces a new valid WordprocessingML document from the listOfComparisonUnitAtoms
                     XDocument newXDoc1 = ProduceNewXDocFromCorrelatedSequence(wDoc.MainDocumentPart, listOfComparisonUnitAtoms, rootNamespaceAttributes, settings);
            
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

        private static XDocument ProduceNewXDocFromCorrelatedSequence(OpenXmlPart part,
            IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList,
            List<XAttribute> rootNamespaceDeclarations,
            WmlComparerSettings settings)
        {
            // fabricate new MainDocumentPart from correlatedSequence

            if (s_DumpLog)
            {
                //dump out content atoms
                var sb = new StringBuilder();
                foreach (var item in comparisonUnitAtomList)
                    sb.Append(item.ToString()).Append(Environment.NewLine);
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }

            s_MaxId = 0;
            XDocument newXDoc = new XDocument();
            var newBodyChildren = CoalesceRecurse(part, comparisonUnitAtomList, 0, settings);
            newXDoc.Add(
                new XElement(W.document,
                    rootNamespaceDeclarations,
                    new XElement(W.body, newBodyChildren)));
            FixUpRevMarkIds(newXDoc);
            FixUpDocPrIds(newXDoc);

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

        private static void FixUpDocPrIds(XDocument newXDoc)
        {
            var docPrToChange = newXDoc
                .Descendants()
                .Where(d => d.Name == WP.docPr);
            var nextId = 1;
            foreach (var item in docPrToChange)
            {
                var idAtt = item.Attribute("id");
                if (idAtt != null)
                    idAtt.Value = (nextId++).ToString();
            }
        }

        private static void FixUpRevMarkIds(XDocument newXDoc)
        {
            var revMarksToChange = newXDoc
                .Descendants()
                .Where(d => d.Name == W.ins || d.Name == W.del);
            var nextId = 0;
            foreach (var item in revMarksToChange)
            {
                var idAtt = item.Attribute(W.id);
                if (idAtt != null)
                    idAtt.Value = (nextId++).ToString();
            }
        }

        private static object CoalesceRecurse(OpenXmlPart part, IEnumerable<ComparisonUnitAtom> list, int level, WmlComparerSettings settings)
        {
            var grouped = list
                .GroupBy(ca =>
                {
                    // per the algorithm, The following condition will never evaluate to true
                    // if it evaluates to true, then the basic mechanism for breaking a hierarchical structure into flat and back is broken.

                    if (level >= ca.AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 2 - why do we have ComparisonUnitAtom objects with fewer ancestors than its siblings?");

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
                var sbs = sb.ToString();
            }

            var elementList = grouped
                .Select(g =>
                {
                    // see the comment above at the beginning of CoalesceRecurse
                    if (level >= g.First().AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 1 - why do we have ComparisonUnitAtom objects with fewer ancestors than its siblings?");

                    var ancestorBeingConstructed = g.First().AncestorElements[level];

                    if (ancestorBeingConstructed.Name == W.p)
                    {
                        var groupedChildren = g
                            .GroupAdjacent(gc => gc.ContentElement.Name.ToString() + " | " + gc.CorrelationStatus.ToString());
                        var newChildElements = groupedChildren
                            .Where(gc => gc.First().ContentElement.Name != W.pPr)
                            .Select(gc =>
                            {
                                return gc.Select(gcc =>
                                {
                                    if (gcc.ContentElement.Name == M.oMath ||
                                        gcc.ContentElement.Name == M.oMathPara)
                                    {
                                        var deleting = gcc.CorrelationStatus == CorrelationStatus.Deleted;
                                        var inserting = gcc.CorrelationStatus == CorrelationStatus.Inserted;

                                        if (deleting)
                                        {
                                            return new XElement(W.del,
                                                new XAttribute(W.author, settings.AuthorForRevisions),
                                                new XAttribute(W.id, s_MaxId++),
                                                new XAttribute(W.date, settings.DateTimeForRevisions),
                                                gcc.ContentElement);
                                        }
                                        else if (inserting)
                                        {
                                            return new XElement(W.ins,
                                                new XAttribute(W.author, settings.AuthorForRevisions),
                                                new XAttribute(W.id, s_MaxId++),
                                                new XAttribute(W.date, settings.DateTimeForRevisions),
                                                gcc.ContentElement);
                                        }
                                        else
                                        {
                                            return gcc.ContentElement;
                                        }
                                    }
                                    else
                                        return CoalesceRecurse(part, new[] { gcc }, level + 1, settings);
                                });
                            });

                        XElement pPr = null;
                        ComparisonUnitAtom pPrComparisonUnitAtom = null;
                        var newParaPropsGroup = groupedChildren
                            .Where(gc => gc.First().ContentElement.Name == W.pPr)
                            .ToList();

                        if (newParaPropsGroup.Any())
                        {
                            pPrComparisonUnitAtom = newParaPropsGroup.First().FirstOrDefault();
                            if (pPrComparisonUnitAtom != null)
                            {
                                pPr = new XElement(pPrComparisonUnitAtom.ContentElement); // clone so we can change it
                                if (pPrComparisonUnitAtom.CorrelationStatus == CorrelationStatus.Deleted)
                                    pPr.Elements(W.sectPr).Remove(); // for now, don't move sectPr from old document to new document.
                            }
                        }
                        if (pPrComparisonUnitAtom != null)
                        {
                            if (pPr == null)
                                pPr = new XElement(W.pPr);
                            // if there are no para props in the group, then don't need to do anything.
                            // if there is one para prop in the group, then may need to mark the paragraph as inserted or deleted.
                            // if there are two para props in the group, then one was inserted, another was deleted, so leave
                            //   pPr alone.
                            if (newParaPropsGroup.Count() == 1)
                            {
                                if (pPrComparisonUnitAtom.CorrelationStatus == CorrelationStatus.Deleted)
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
                                else if (pPrComparisonUnitAtom.CorrelationStatus == CorrelationStatus.Inserted)
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
                    if (ancestorBeingConstructed.Name == W.hyperlink)
                        return ReconstructElement(part, g, ancestorBeingConstructed, null, null, level, settings);
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
                .Where(d => d.Attributes().Any(a => ComparisonUnitWord.s_RelationshipAttributeNames.Contains(a.Name)))
                .ToList();
            foreach (var element in elementsToUpdate)
            {
                var attributesToUpdate = element
                    .Attributes()
                    .Where(a => ComparisonUnitWord.s_RelationshipAttributeNames.Contains(a.Name))
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

        private static XAttribute GetXmlSpaceAttribute(string textOfTextElement)
        {
            if (char.IsWhiteSpace(textOfTextElement[0]) ||
                char.IsWhiteSpace(textOfTextElement[textOfTextElement.Length - 1]))
                return new XAttribute(XNamespace.Xml + "space", "preserve");
            return null;
        }

        private static XElement ReconstructElement(OpenXmlPart part, IGrouping<string, ComparisonUnitAtom> g, XElement ancestorBeingConstructed, XName props1XName,
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

        private static List<CorrelatedSequence> Lcs(ComparisonUnit[] cu1, ComparisonUnit[] cu2, WmlComparerSettings settings)
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

            while (true)
            {
                if (s_DumpLog)
                {
                    var sb = new StringBuilder();
                    foreach (var item in csList)
                        sb.Append(item.ToString()).Append(Environment.NewLine);
                    var sbs = sb.ToString();
                    //TestUtil.NotePad(sbs);
                    Console.WriteLine(sbs);
                }

                var unknown = csList
                    .FirstOrDefault(z => z.CorrelationStatus == CorrelationStatus.Unknown);
                if (unknown != null)
                {
                    if (s_DumpLog)
                    {
                        var sb = new StringBuilder();
                        sb.Append(unknown.ToString());
                        var sbs = sb.ToString();
                        Console.WriteLine(sbs);
                    }

                    var newSequence = FindCommonAtBeginningAndEnd(unknown, settings);
                    if (newSequence == null)
                    {
                        newSequence = DoLcsAlgorithm(unknown, settings);
                    }

                    var indexOfUnknown = csList.IndexOf(unknown);
                    csList.Remove(unknown);

                    newSequence.Reverse();
                    foreach (var item in newSequence)
                        csList.Insert(indexOfUnknown, item);

                    continue;
                }
                return csList;
            }
        }

        private static List<CorrelatedSequence> DoLcsAlgorithm(CorrelatedSequence unknown, WmlComparerSettings settings)
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

            // never start a common section with a paragraph mark.
            while (true)
            {
                if (currentLongestCommonSequenceLength <= 1)
                    break;

                var firstCommon = cul1[currentI1];

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                if (firstCommonWord == null)
                    break;

                // if the word contains more than one atom, then not a paragraph mark
                if (firstCommonWord.Contents.Count() != 1)
                    break;

                var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                if (firstCommonAtom == null)
                    break;

                if (firstCommonAtom.ContentElement.Name != W.pPr)
                    break;

                --currentLongestCommonSequenceLength;
                if (currentLongestCommonSequenceLength == 0)
                {
                    currentI1 = -1;
                    currentI2 = -1;
                }
                else
                {
                    ++currentI1;
                    ++currentI2;
                }
            }

            bool isOnlyParagraphMark = false;
            if (currentLongestCommonSequenceLength == 1)
            {
                var firstCommon = cul1[currentI1];

                var firstCommonWord = firstCommon as ComparisonUnitWord;
                if (firstCommonWord != null)
                {
                    // if the word contains more than one atom, then not a paragraph mark
                    if (firstCommonWord.Contents.Count() == 1)
                    {
                        var firstCommonAtom = firstCommonWord.Contents.First() as ComparisonUnitAtom;
                        if (firstCommonAtom != null)
                        {
                            if (firstCommonAtom.ContentElement.Name == W.pPr)
                                isOnlyParagraphMark = true;
                        }
                    }
                }
            }

            // don't match just a single space
            if (currentLongestCommonSequenceLength == 1)
            {
                var cuw2 = cul2[currentI2] as ComparisonUnitAtom;
                if (cuw2 != null)
                {
                    if (cuw2.ContentElement.Name == W.t && cuw2.ContentElement.Value == " ")
                    {
                        currentI1 = -1;
                        currentI2 = -1;
                        currentLongestCommonSequenceLength = 0;
                    }
                }
            }

            // if we are only looking at text, and if the longest common subsequence is less than 15% of the whole, then forget it,
            // don't find that LCS.
            if (!isOnlyParagraphMark && currentLongestCommonSequenceLength > 0)
            {
                var anyButWord1 = cul1.Any(cu => (cu as ComparisonUnitWord) == null);
                var anyButWord2 = cul2.Any(cu => (cu as ComparisonUnitWord) == null);
                if (!anyButWord1 && !anyButWord2)
                {
                    var maxLen = Math.Max(cul1.Length, cul2.Length);
                    if (((double)currentLongestCommonSequenceLength / (double)maxLen) < settings.DetailThreshold)
                    {
                        currentI1 = -1;
                        currentI2 = -1;
                        currentLongestCommonSequenceLength = 0;
                    }
                }
            }

            var newListOfCorrelatedSequence = new List<CorrelatedSequence>();
            if (currentI1 == -1 && currentI2 == -1)
            {
                var leftLength = unknown.ComparisonUnitArray1.Length;
                var leftTables = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Table).Count();
                var leftRows = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Row).Count();
                var leftCells = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell).Count();
                var leftParagraphs = unknown.ComparisonUnitArray1.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph).Count();
                var leftWords = unknown.ComparisonUnitArray1.OfType<ComparisonUnitWord>().Count();

                var rightLength = unknown.ComparisonUnitArray2.Length;
                var rightTables = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Table).Count();
                var rightRows = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Row).Count();
                var rightCells = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell).Count();
                var rightParagraphs = unknown.ComparisonUnitArray2.OfType<ComparisonUnitGroup>().Where(l => l.ComparisonUnitGroupType == ComparisonUnitGroupType.Paragraph).Count();
                var rightWords = unknown.ComparisonUnitArray2.OfType<ComparisonUnitWord>().Count();

                // if either side has both words and rows, then we need to separate out into separate unknown correlated sequences
                // group adjacent based on whether word or row
                // in most cases, the count of groups will be the same, but they may differ
                // if the first group on either side is word, then create a deleted or inserted corr sequ for it.
                // then have counter on both sides pointing to the first matched pairs of rows
                // create an unknown corr sequ for it.
                // increment both counters
                // if one is at end but the other is not, then tag the remaining content as inserted or deleted, and done.
                // if both are at the end, then done
                // return the new list of corr sequ

                var leftOnlyWordsAndRows = leftLength == leftWords + leftRows;
                var rightOnlyWordsAndRows = rightLength == rightWords + rightRows;
                if ((leftWords > 0 && rightWords > 0) &&
                    (leftRows > 0 || rightRows > 0) &&
                    (leftOnlyWordsAndRows && rightOnlyWordsAndRows))
                {

                    var leftGrouped = unknown
                        .ComparisonUnitArray1
                        .GroupAdjacent(cu =>
                        {
                            if (cu is ComparisonUnitWord)
                                return "Word";
                            else
                                return "Row";
                        })
                        .ToArray();
                    var rightGrouped = unknown
                        .ComparisonUnitArray2
                        .GroupAdjacent(cu =>
                        {
                            if (cu is ComparisonUnitWord)
                                return "Word";
                            else
                                return "Row";
                        })
                        .ToArray();
                    int iLeft = 0;
                    int iRight = 0;

                    // create an unknown corr sequ for it.
                    // increment both counters
                    // if one is at end but the other is not, then tag the remaining content as inserted or deleted, and done.
                    // if both are at the end, then done
                    // return the new list of corr sequ

                    while (true)
                    {
                        if ((leftGrouped[iLeft].Key == "Word" && rightGrouped[iRight].Key == "Word") ||
                            (leftGrouped[iLeft].Key == "Row" && rightGrouped[iRight].Key == "Row"))
                        {
                            var unknownCorrelatedSequence = new CorrelatedSequence();
                            unknownCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[iLeft].ToArray();
                            unknownCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
                            unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
                            ++iLeft;
                            ++iRight;
                        }
                        else if (leftGrouped[iLeft].Key == "Word" && rightGrouped[iRight].Key == "Row")
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence();
                            deletedCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[iLeft].ToArray();
                            deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                            deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            ++iLeft;
                        }
                        else if (leftGrouped[iLeft].Key == "Row" && rightGrouped[iRight].Key == "Word")
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[iRight].ToArray();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            ++iRight;
                        }

                        if (iLeft == leftGrouped.Length && iRight == rightGrouped.Length)
                            return newListOfCorrelatedSequence;

                        // if there is content on the left, but not content on the right
                        if (iRight == rightGrouped.Length)
                        {
                            for (int j = iLeft; j < leftGrouped.Length; j++)
                            {
                                var deletedCorrelatedSequence = new CorrelatedSequence();
                                deletedCorrelatedSequence.ComparisonUnitArray1 = leftGrouped[j].ToArray();
                                deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                            }
                            return newListOfCorrelatedSequence;
                        }
                        // there is content on the right but not on the left
                        else if (iLeft == leftGrouped.Length) 
                        {
                            for (int j = iRight; j < rightGrouped.Length; j++)
                            {
                                var insertedCorrelatedSequence = new CorrelatedSequence();
                                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                                insertedCorrelatedSequence.ComparisonUnitArray2 = rightGrouped[j].ToArray();
                                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                            }
                            return newListOfCorrelatedSequence;
                        }
                        // else continue on next round.
                    }
                }

                // If either side contains only paras or tables, then flatten and iterate.
                var leftOnlyParasAndTables = leftLength == leftTables + leftParagraphs;
                var rightOnlyParasAndTables = rightLength == rightTables + rightParagraphs;
                if (leftOnlyParasAndTables && rightOnlyParasAndTables)
                {
                    // flatten paras and tables, and iterate
                    var left = unknown
                        .ComparisonUnitArray1
                        .Select(cu => cu.Contents)
                        .SelectMany(m => m)
                        .ToArray();

                    var right = unknown
                        .ComparisonUnitArray2
                        .Select(cu => cu.Contents)
                        .SelectMany(m => m)
                        .ToArray();

                    var unknownCorrelatedSequence = new CorrelatedSequence();
                    unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                    unknownCorrelatedSequence.ComparisonUnitArray1 = left;
                    unknownCorrelatedSequence.ComparisonUnitArray2 = right;
                    newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);

                    return newListOfCorrelatedSequence;
                }

                // if first of left is a row and first of right is a row
                // then flatten the row to cells and iterate.

                var firstLeft = unknown
                    .ComparisonUnitArray1
                    .First() as ComparisonUnitGroup;

                var firstRight = unknown
                    .ComparisonUnitArray2
                    .First() as ComparisonUnitGroup;

                if (firstLeft != null && firstRight != null)
                {
                    if (firstLeft.ComparisonUnitGroupType == ComparisonUnitGroupType.Row &&
                        firstRight.ComparisonUnitGroupType == ComparisonUnitGroupType.Row)
                    {
                        ComparisonUnit[] leftContent = firstLeft.Contents.ToArray();
                        ComparisonUnit[] rightContent = firstRight.Contents.ToArray();

                        var lenLeft = leftContent.Length;
                        var lenRight = rightContent.Length;

                        if (lenLeft < lenRight)
                            leftContent = leftContent.Concat(Enumerable.Repeat<ComparisonUnit>(null, lenRight - lenLeft)).ToArray();
                        else if (lenRight < lenLeft)
                            rightContent = rightContent.Concat(Enumerable.Repeat<ComparisonUnit>(null, lenLeft - lenRight)).ToArray();

                        List<CorrelatedSequence> newCs = leftContent.Zip(rightContent, (l, r) =>
                            {
                                if (l != null && r != null)
                                {
                                    var cellLcs = Lcs(l.Contents.ToArray(), r.Contents.ToArray(), settings);
                                    return cellLcs.ToArray();
                                }
                                if (l == null)
                                {
                                    var insertedCorrelatedSequence = new CorrelatedSequence();
                                    insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                                    insertedCorrelatedSequence.ComparisonUnitArray2 = r.Contents.ToArray();
                                    insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                                    return new[] { insertedCorrelatedSequence };
                                }
                                else if (r == null)
                                {
                                    var deletedCorrelatedSequence = new CorrelatedSequence();
                                    deletedCorrelatedSequence.ComparisonUnitArray1 = l.Contents.ToArray();
                                    deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                                    deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                                    return new[] { deletedCorrelatedSequence };
                                }
                                else
                                    throw new OpenXmlPowerToolsException("Internal error");
                            })
                            .SelectMany(m => m)
                            .ToList();

                        foreach (var cs in newCs)
                            newListOfCorrelatedSequence.Add(cs);

                        var remainderLeft = unknown
                            .ComparisonUnitArray1
                            .Skip(1)
                            .ToArray();

                        var remainderRight = unknown
                            .ComparisonUnitArray2
                            .Skip(1)
                            .ToArray();

                        if (remainderLeft.Length > 0 && remainderRight.Length == 0)
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence();
                            deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                            deletedCorrelatedSequence.ComparisonUnitArray1 = remainderLeft;
                            deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                        }
                        else if (remainderRight.Length > 0 && remainderLeft.Length == 0)
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = remainderRight;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                        }
                        else if (remainderLeft.Length > 0 && remainderRight.Length > 0)
                        {
                            var unknownCorrelatedSequence2 = new CorrelatedSequence();
                            unknownCorrelatedSequence2.CorrelationStatus = CorrelationStatus.Unknown;
                            unknownCorrelatedSequence2.ComparisonUnitArray1 = remainderLeft;
                            unknownCorrelatedSequence2.ComparisonUnitArray2 = remainderRight;
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence2);
                        }

                        if (s_DumpLog)
                        {
                            var sb = new StringBuilder();
                            foreach (var item in newListOfCorrelatedSequence)
                                sb.Append(item.ToString()).Append(Environment.NewLine);
                            var sbs = sb.ToString();
                            Console.WriteLine(sbs);
                        }

                        return newListOfCorrelatedSequence;
                    }
                    if (firstLeft.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell &&
                        firstRight.ComparisonUnitGroupType == ComparisonUnitGroupType.Cell)
                    {
                        var left = firstLeft
                            .Contents
                            .ToArray();

                        var right = firstRight
                            .Contents
                            .ToArray();

                        var unknownCorrelatedSequence = new CorrelatedSequence();
                        unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                        unknownCorrelatedSequence.ComparisonUnitArray1 = left;
                        unknownCorrelatedSequence.ComparisonUnitArray2 = right;
                        newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);

                        var remainderLeft = unknown
                            .ComparisonUnitArray1
                            .Skip(1)
                            .ToArray();

                        var remainderRight = unknown
                            .ComparisonUnitArray2
                            .Skip(1)
                            .ToArray();

                        if (remainderLeft.Length > 0 && remainderRight.Length == 0)
                        {
                            var deletedCorrelatedSequence = new CorrelatedSequence();
                            deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                            deletedCorrelatedSequence.ComparisonUnitArray1 = remainderLeft;
                            deletedCorrelatedSequence.ComparisonUnitArray2 = null;
                            newListOfCorrelatedSequence.Add(deletedCorrelatedSequence);
                        }
                        else if (remainderRight.Length > 0 && remainderLeft.Length == 0)
                        {
                            var insertedCorrelatedSequence = new CorrelatedSequence();
                            insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                            insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                            insertedCorrelatedSequence.ComparisonUnitArray2 = remainderRight;
                            newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
                        }
                        else if (remainderLeft.Length > 0 && remainderRight.Length > 0)
                        {
                            var unknownCorrelatedSequence2 = new CorrelatedSequence();
                            unknownCorrelatedSequence2.CorrelationStatus = CorrelationStatus.Unknown;
                            unknownCorrelatedSequence2.ComparisonUnitArray1 = remainderLeft;
                            unknownCorrelatedSequence2.ComparisonUnitArray2 = remainderRight;
                            newListOfCorrelatedSequence.Add(unknownCorrelatedSequence2);
                        }

                        return newListOfCorrelatedSequence;
                    }
                }

                // otherwise create ins and del

                var deletedCorrelatedSequence3 = new CorrelatedSequence();
                deletedCorrelatedSequence3.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence3.ComparisonUnitArray1 = unknown.ComparisonUnitArray1;
                deletedCorrelatedSequence3.ComparisonUnitArray2 = null;
                newListOfCorrelatedSequence.Add(deletedCorrelatedSequence3);

                var insertedCorrelatedSequence3 = new CorrelatedSequence();
                insertedCorrelatedSequence3.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence3.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence3.ComparisonUnitArray2 = unknown.ComparisonUnitArray2;
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence3);

                return newListOfCorrelatedSequence;
            }

            if (currentI1 > 0 && currentI2 == 0)
            {
                var deletedCorrelatedSequence = new CorrelatedSequence();
                deletedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Deleted;
                deletedCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Take(currentI1)
                    .ToArray();
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
                    .ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (currentI1 > 0 && currentI2 > 0)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Take(currentI1)
                    .ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Take(currentI2)
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
                .ToArray();
            middleEqual.ComparisonUnitArray2 = cul2
                .Skip(currentI2)
                .Take(currentLongestCommonSequenceLength)
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
                    .ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (endI1 < cul1.Length && endI2 < cul2.Length)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1
                    .Skip(endI1)
                    .ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2
                    .Skip(endI2)
                    .ToArray();
                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (endI1 == cul1.Length && endI2 == cul2.Length)
            {
                // nothing to do
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

        private class Atgbw
        {
            public int? Key;
            public ComparisonUnitAtom ComparisonUnitAtomMember;
            public int NextIndex;
        }

        private static ComparisonUnit[] GetComparisonUnitList(ComparisonUnitAtom[] comparisonUnitAtomList, WmlComparerSettings settings)
        {
            var seed = new Atgbw()
            {
                Key = null,
                ComparisonUnitAtomMember = null,
                NextIndex = 0,
            };

            var groupingKey = comparisonUnitAtomList
                .Rollup(seed, (sr, prevAtgbw, i) =>
                {
                    int? key = null;
                    var nextIndex = prevAtgbw.NextIndex;
                    if (sr.ContentElement.Name == W.t)
                    {
                        string chr = sr.ContentElement.Value;
                        var ch = chr[0];
                        if (ch == '.' || ch == ',')
                        {
                            bool beforeIsDigit = false;
                            if (i > 0)
                            {
                                var prev = comparisonUnitAtomList[i - 1];
                                if (prev.ContentElement.Name == W.t && char.IsDigit(prev.ContentElement.Value[0]))
                                    beforeIsDigit = true;
                            }
                            bool afterIsDigit = false;
                            if (i < comparisonUnitAtomList.Length - 1)
                            {
                                var next = comparisonUnitAtomList[i + 1];
                                if (next.ContentElement.Name == W.t && char.IsDigit(next.ContentElement.Value[0]))
                                    afterIsDigit = true;
                            }
                            if (beforeIsDigit || afterIsDigit)
                            {
                                key = nextIndex;
                            }
                            else
                            {
                                nextIndex++;
                                key = nextIndex;
                                nextIndex++;
                            }
                        }
                        else if (settings.WordSeparators.Contains(ch))
                        {
                            nextIndex++;
                            key = nextIndex;
                            nextIndex++;
                        }
                        else
                        {
                            key = nextIndex;
                        }
                    }
                    else if (WordBreakElements.Contains(sr.ContentElement.Name))
                    {
                        nextIndex++;
                        key = nextIndex;
                        nextIndex++;
                    }
                    else
                    {
                        key = nextIndex;
                    }
                    return new Atgbw()
                    {
                        Key = key,
                        ComparisonUnitAtomMember = sr,
                        NextIndex = nextIndex,
                    };
                });

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var item in groupingKey)
                {
                    sb.Append(item.Key + Environment.NewLine);
                    sb.Append("    " + item.ComparisonUnitAtomMember.ToString(0) + Environment.NewLine);
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
                        sb.Append("    " + gc.ComparisonUnitAtomMember.ToString(0) + Environment.NewLine);
                    }
                }
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }

            var withHierarchicalGroupingKey = groupedByWords
                .Select(g =>
                    {
                        var hierarchicalGroupingArray = g
                            .First()
                            .ComparisonUnitAtomMember
                            .AncestorElements
                            .Where(a => ComparisonGroupingElements.Contains(a.Name))
                            .Select(a => a.Name.LocalName + ":" + (string)a.Attribute(PtOpenXml.Unid))
                            .ToArray();

                        return new WithHierarchicalGroupingKey() {
                            ComparisonUnitWord = new ComparisonUnitWord(g.Select(gc => gc.ComparisonUnitAtomMember)),
                            HierarchicalGroupingArray = hierarchicalGroupingArray,
                        };
                    }
                )
                .ToArray();

            if (s_DumpLog)
            {
                var sb = new StringBuilder();
                foreach (var group in withHierarchicalGroupingKey)
                {
                    sb.Append("Grouping Array: " + group.HierarchicalGroupingArray.Select(gam => gam + " - ").StringConcatenate() + Environment.NewLine);
                    foreach (var gc in group.ComparisonUnitWord.Contents)
                    {
                        sb.Append("    " + gc.ToString(0) + Environment.NewLine);
                    }
                }
                var sbs = sb.ToString();
                Console.WriteLine(sbs);
            }

            var cul = GetHierarchicalComparisonUnits(withHierarchicalGroupingKey, 0).ToArray();

            if (s_DumpLog)
            {
                var str = ComparisonUnit.ComparisonUnitListToString(cul);
                Console.WriteLine(str);
            }

            return cul;
        }

        private static IEnumerable<ComparisonUnit> GetHierarchicalComparisonUnits(IEnumerable<WithHierarchicalGroupingKey> input, int level)
        {
            var grouped = input
                .GroupAdjacent(whgk =>
                {
                    if (level >= whgk.HierarchicalGroupingArray.Length)
                        return "";
                    return whgk.HierarchicalGroupingArray[level];
                });
            var retList = grouped
                .Select(gc =>
                {
                    if (gc.Key == "")
                    {
                        return (IEnumerable<ComparisonUnit>)gc.Select(whgk => whgk.ComparisonUnitWord).ToList();
                    }
                    else
                    {
                        ComparisonUnitGroupType? group = null;
                        var spl = gc.Key.Split(':');
                        if (spl[0] == "p")
                            group = ComparisonUnitGroupType.Paragraph;
                        else if (spl[0] == "tbl")
                            group = ComparisonUnitGroupType.Table;
                        else if (spl[0] == "tr")
                            group = ComparisonUnitGroupType.Row;
                        else if (spl[0] == "tc")
                            group = ComparisonUnitGroupType.Cell;
                        var newCompUnitGroup = new ComparisonUnitGroup(GetHierarchicalComparisonUnits(gc, level + 1), (ComparisonUnitGroupType)group);
                        return new[] { newCompUnitGroup };
                    }
                })
                .SelectMany(m => m)
                .ToList();
            return retList;
        }

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
            W.commentRangeStart,
            W.commentRangeEnd,
            W.lastRenderedPageBreak,
            W.proofErr,
            W.tblPr,
            W.sectPr,
            W.permEnd,
            W.permStart,
        };

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

        private static XName[] ElementsToHaveSha1Hash = new XName[]
        {
            W.p,
            W.tbl,
            W.tr,
            W.tc,
            W.drawing,
        };

        private static XName[] InvalidElements = new XName[]
        {
            W.altChunk,
            W.customXml,
            W.customXmlDelRangeEnd,
            W.customXmlDelRangeStart,
            W.customXmlInsRangeEnd,
            W.customXmlInsRangeStart,
            W.customXmlMoveFromRangeEnd,
            W.customXmlMoveFromRangeStart,
            W.customXmlMoveToRangeEnd,
            W.customXmlMoveToRangeStart,
            W.moveFrom,
            W.moveFromRangeStart,
            W.moveFromRangeEnd,
            W.moveTo,
            W.moveToRangeStart,
            W.moveToRangeEnd,
            W.subDoc,
        };

        private class RecursionInfo
        {
            public XName ElementName;
            public XName[] ChildElementPropertyNames;
        }

        private static RecursionInfo[] RecursionElements = new RecursionInfo[]
        {
            new RecursionInfo()
            {
                ElementName = W.del,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.ins,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.tbl,
                ChildElementPropertyNames = new[] { W.tblPr, W.tblGrid, W.tblPrEx },
            },
            new RecursionInfo()
            {
                ElementName = W.tr,
                ChildElementPropertyNames = new[] { W.trPr, W.tblPrEx },
            },
            new RecursionInfo()
            {
                ElementName = W.tc,
                ChildElementPropertyNames = new[] { W.tcPr, W.tblPrEx },
            },
            new RecursionInfo()
            {
                ElementName = W.sdt,
                ChildElementPropertyNames = new[] { W.sdtPr, W.sdtEndPr },
            },
            new RecursionInfo()
            {
                ElementName = W.sdtContent,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.hyperlink,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.fldSimple,
                ChildElementPropertyNames = null,
            },
            new RecursionInfo()
            {
                ElementName = W.smartTag,
                ChildElementPropertyNames = new[] { W.smartTagPr },
            },
        };

        internal static List<ComparisonUnitAtom> CreateComparisonUnitAtomList(WordprocessingDocument wDoc, OpenXmlPart part)
        {
            VerifyNoInvalidContent(part);
            AssignIdToAllElements(part);  // add the Guid id to every element for which we need to establish identity
            MoveLastSectPrIntoLastParagraph(part);
            var cal = CreateComparisonUnitAtomListInternal(part);
            return cal;
        }

        private static void VerifyNoInvalidContent(OpenXmlPart part)
        {
            var xDoc = part.GetXDocument();
            var invalidElement = xDoc.Descendants().FirstOrDefault(d => InvalidElements.Contains(d.Name));
            if (invalidElement == null)
                return;
            throw new NotSupportedException("Document contains " + invalidElement.Name.LocalName);
        }

        internal static XDocument Coalesce(List<ComparisonUnitAtom> comparisonUnitAtomList)
        {
            XDocument newXDoc = new XDocument();
            var newBodyChildren = CoalesceRecurse(comparisonUnitAtomList, 0);
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

        private static object CoalesceRecurse(IEnumerable<ComparisonUnitAtom> list, int level)
        {
            var grouped = list
                .GroupBy(sr =>
                {
                    // per the algorithm, The following condition will never evaluate to true
                    // if it evaluates to true, then the basic mechanism for breaking a hierarchical structure into flat and back is broken.

                    // for a table, we initially get all ComparisonUnitAtoms for the entire table, then process.  When processing a row,
                    // no ComparisonUnitAtoms will have ancestors outside the row.  Ditto for cells, and on down the tree.
                    if (level >= sr.AncestorElements.Length)
                        throw new OpenXmlPowerToolsException("Internal error 4 - why do we have ComparisonUnitAtom objects with fewer ancestors than its siblings?");

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
                        throw new OpenXmlPowerToolsException("Internal error 3 - why do we have ComparisonUnitAtom objects with fewer ancestors than its siblings?");

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
                                var name = gc.First().ContentElement.Name;
                                if (name == W.t || name == W.delText)
                                {
                                    var textOfTextElement = gc.Select(gce => gce.ContentElement.Value).StringConcatenate();
                                    return (object)(new XElement(name,
                                        GetXmlSpaceAttribute(textOfTextElement),
                                        textOfTextElement));
                                }
                                else
                                    return gc.Select(gce => gce.ContentElement);
                            });
                        var runProps = ancestorBeingConstructed.Elements(W.rPr);
                        return new XElement(W.r, runProps, newChildElements);
                    }

                    var re = RecursionElements.FirstOrDefault(z => z.ElementName == ancestorBeingConstructed.Name);
                    if (re != null)
                    {
                        return ReconstructElement(g, ancestorBeingConstructed, re.ChildElementPropertyNames, level);
                    }

                    var newElement = new XElement(ancestorBeingConstructed.Name,
                        ancestorBeingConstructed.Attributes(),
                        CoalesceRecurse(g, level + 1));
                    return newElement;
                })
                .ToList();
            return elementList;
        }

        private static XElement ReconstructElement(IGrouping<string, ComparisonUnitAtom> g, XElement ancestorBeingConstructed, XName[] childPropElementNames, int level)
        {
            var newChildElements = CoalesceRecurse(g, level + 1);
            IEnumerable<XElement> childProps = null;
            if (childPropElementNames != null)
                childProps = ancestorBeingConstructed.Elements()
                    .Where(a => childPropElementNames.Contains(a.Name));

            var reconstructedElement = new XElement(ancestorBeingConstructed.Name, childProps, newChildElements);
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

        private static List<ComparisonUnitAtom> CreateComparisonUnitAtomListInternal(OpenXmlPart part)
        {
            var partXDoc = part.GetXDocument();
            XElement root = null;
            if (part is MainDocumentPart)
                root = partXDoc.Root.Element(W.body);
            else
                root = partXDoc.Root;

            var comparisonUnitAtomList = new List<ComparisonUnitAtom>();
            CreateComparisonUnitAtomListRecurse(part, root, comparisonUnitAtomList);
            return comparisonUnitAtomList;
        }

        private static XName[] ComparisonGroupingElements = new[] {
            W.p,
            W.tbl,
            W.tr,
            W.tc,
        };

        private static void CreateComparisonUnitAtomListRecurse(OpenXmlPart part, XElement element, List<ComparisonUnitAtom> comparisonUnitAtomList)
        {
            if (element.Name == W.body)
            {
                foreach (var item in element.Elements())
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList);
                return;
            }

            if (element.Name == W.p)
            {
                var paraChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.pPr);
                foreach (var item in paraChildrenToProcess)
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList);
                var paraProps = element.Element(W.pPr);
                if (paraProps == null)
                {
                    ComparisonUnitAtom pPrComparisonUnitAtom = new ComparisonUnitAtom(
                        new XElement(W.pPr),
                        element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body).Reverse().ToArray(),
                        part);
                    comparisonUnitAtomList.Add(pPrComparisonUnitAtom);
                }
                else
                {
                    ComparisonUnitAtom pPrComparisonUnitAtom = new ComparisonUnitAtom(
                        paraProps,
                        element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body).Reverse().ToArray(),
                        part);
                    comparisonUnitAtomList.Add(pPrComparisonUnitAtom);
                }
                return;
            }

            if (element.Name == W.r)
            {
                var runChildrenToProcess = element
                    .Elements()
                    .Where(e => e.Name != W.rPr);
                foreach (var item in runChildrenToProcess)
                    CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList);
                return;
            }

            if (element.Name == W.t || element.Name == W.delText)
            {
                var val = element.Value;
                foreach (var ch in val)
                {
                    ComparisonUnitAtom sr = new ComparisonUnitAtom(
                        new XElement(element.Name, ch),
                        element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body).Reverse().ToArray(),
                        part);
                    comparisonUnitAtomList.Add(sr);
                }
                return;
            }

            if (AllowableRunChildren.Contains(element.Name))
            {
                ComparisonUnitAtom sr3 = new ComparisonUnitAtom(
                    element,
                    element.AncestorsAndSelf().TakeWhile(a => a.Name != W.body).Reverse().ToArray(),
                    part);
                comparisonUnitAtomList.Add(sr3);
                return;
            }

            // todo use recursioninfo array here
            var re = RecursionElements.FirstOrDefault(z => z.ElementName == element.Name);
            if (re != null)
            {
                AnnotateElementWithProps(part, element, comparisonUnitAtomList, re.ChildElementPropertyNames);
                return;
            }

            if (ElementsToThrowAway.Contains(element.Name))
                return;

            throw new OpenXmlPowerToolsException("Internal error - unexpected element");
        }

        private static void AnnotateElementWithProps(OpenXmlPart part, XElement element, List<ComparisonUnitAtom> comparisonUnitAtomList, XName[] childElementPropertyNames)
        {
            IEnumerable<XElement> runChildrenToProcess = null;
            if (childElementPropertyNames == null)
                runChildrenToProcess = element.Elements();
            else
                runChildrenToProcess = element
                    .Elements()
                    .Where(e => !childElementPropertyNames.Contains(e.Name));

            foreach (var item in runChildrenToProcess)
                CreateComparisonUnitAtomListRecurse(part, item, comparisonUnitAtomList);
        }

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

    internal class WithHierarchicalGroupingKey
    {
        public string[] HierarchicalGroupingArray;
        public ComparisonUnitWord ComparisonUnitWord;
    }

    public abstract class ComparisonUnit
    {
        public List<ComparisonUnit> Contents;
        public string SHA1Hash;
        public CorrelationStatus CorrelationStatus;

        public IEnumerable<ComparisonUnit> Descendants()
        {
            List<ComparisonUnit> comparisonUnitList = new List<ComparisonUnit>();
            DescendantsInternal(this, comparisonUnitList);
            return comparisonUnitList;
        }

        public IEnumerable<ComparisonUnitAtom> DescendantContentAtoms()
        {
            return Descendants().OfType<ComparisonUnitAtom>();
        }

        private void DescendantsInternal(ComparisonUnit comparisonUnit, List<ComparisonUnit> comparisonUnitList)
        {
            foreach (var cu in comparisonUnit.Contents)
            {
                comparisonUnitList.Add(cu);
                if (cu.Contents != null && cu.Contents.Any())
                    DescendantsInternal(cu, comparisonUnitList);
            }
        }

        public abstract string ToString(int indent);

        internal static object ComparisonUnitListToString(ComparisonUnit[] cul)
        {
            var sb = new StringBuilder();
            sb.Append("Dump Comparision Unit List To String" + Environment.NewLine);
            foreach (var item in cul)
            {
                sb.Append(item.ToString(2) + Environment.NewLine);
            }
            return sb.ToString();
        }
    }

    internal class ComparisonUnitWord : ComparisonUnit
    {
        public ComparisonUnitWord(IEnumerable<ComparisonUnitAtom> comparisonUnitAtomList)
        {
            Contents = comparisonUnitAtomList.OfType<ComparisonUnit>().ToList();
            var sha1String = Contents
                .Select(c => c.SHA1Hash)
                .StringConcatenate();
            SHA1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(sha1String);
        }

        public static XName[] s_ElementsWithRelationshipIds = new XName[] {
            A.blip,
            A.hlinkClick,
            A.relIds,
            C.chart,
            C.externalData,
            C.userShapes,
            DGM.relIds,
            O.OLEObject,
            VML.fill,
            VML.imagedata,
            VML.stroke,
            W.altChunk,
            W.attachedTemplate,
            W.control,
            W.dataSource,
            W.embedBold,
            W.embedBoldItalic,
            W.embedItalic,
            W.embedRegular,
            W.footerReference,
            W.headerReference,
            W.headerSource,
            W.hyperlink,
            W.printerSettings,
            W.recipientData,
            W.saveThroughXslt,
            W.sourceFileName,
            W.src,
            W.subDoc,
            WNE.toolbarData,
        };

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

        public override string ToString(int indent)
        {
            var sb = new StringBuilder();
            sb.Append("".PadRight(indent) + "Word SHA1:" + this.SHA1Hash + Environment.NewLine);
            foreach (var comparisonUnitAtom in Contents)
                sb.Append(comparisonUnitAtom.ToString(indent + 2) + Environment.NewLine);
            return sb.ToString();
        }
    }

    class WmlComparerUtil
    {
        public static string SHA1HashStringForUTF8String(string s)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(s);
            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            var sha1 = SHA1.Create();
            byte[] hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string HexStringFromBytes(byte[] bytes)
        {
            var sb = new StringBuilder();
            foreach (byte b in bytes)
            {
                var hex = b.ToString("x2");
                sb.Append(hex);
            }
            return sb.ToString();
        }
    }

    public class ComparisonUnitAtom : ComparisonUnit
    {
        // AncestorElements are kept in order from the body to the leaf, because this is the order in which we need to access in order
        // to reassemble the document.  However, in many places in the code, it is necessary to find the nearest ancestor, i.e. cell
        // so it is necessary to reverse the order when looking for it, i.e. look from the leaf back to the body element.

        public XElement[] AncestorElements;
        public XElement ContentElement;
        public OpenXmlPart Part;

        public ComparisonUnitAtom(XElement contentElement, XElement[] ancestorElements, OpenXmlPart part)
        {
            ContentElement = contentElement;
            AncestorElements = ancestorElements;
            Part = part;
            CorrelationStatus = GetCorrelationStatusFromAncestors(AncestorElements);
            string sha1Hash = (string)contentElement.Attribute(PtOpenXml.SHA1Hash);
            if (sha1Hash != null)
            {
                SHA1Hash = sha1Hash;
            }
            else
            {
                var shaHashString = GetSha1HashStringForElement(ContentElement);
                SHA1Hash = WmlComparerUtil.SHA1HashStringForUTF8String(shaHashString);
            }
        }

        private string GetSha1HashStringForElement(XElement contentElement)
        {
            return contentElement.Name.LocalName + contentElement.Value;
        }

        private static CorrelationStatus GetCorrelationStatusFromAncestors(XElement[] ancestors)
        {
            var deleted = ancestors.Any(a => a.Name == W.del);
            var inserted = ancestors.Any(a => a.Name == W.ins);
            if (deleted)
                return CorrelationStatus.Deleted;
            else if (inserted)
                return CorrelationStatus.Inserted;
            else
                return CorrelationStatus.Normal;
        }
        
        public override string ToString(int indent)
        {
            int xNamePad = 16;
            var indentString = "".PadRight(indent);

            var sb = new StringBuilder();
            sb.Append(indentString);
            string correlationStatus = "";
            if (CorrelationStatus != OpenXmlPowerTools.CorrelationStatus.Nil)
                correlationStatus = string.Format("[{0}] ", CorrelationStatus.ToString().PadRight(8));
            if (ContentElement.Name == W.t || ContentElement.Name == W.delText)
            {
                sb.AppendFormat("Atom {0}: {1} {2} SHA1:{3} ", PadLocalName(xNamePad, this), ContentElement.Value, correlationStatus, this.SHA1Hash);
                AppendAncestorsDump(sb, this);
            }
            else
            {
                sb.AppendFormat("Atom {0}:   {1} SHA1:{2} ", PadLocalName(xNamePad, this), correlationStatus, this.SHA1Hash);
                AppendAncestorsDump(sb, this);
            }
            return sb.ToString();
        }

        public override string ToString()
        {
            return ToString(0);
        }

        private static string PadLocalName(int xNamePad, ComparisonUnitAtom item)
        {
            return (item.ContentElement.Name.LocalName + " ").PadRight(xNamePad, '-') + " ";
        }

        private void AppendAncestorsDump(StringBuilder sb, ComparisonUnitAtom sr)
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

        public static string ComparisonUnitAtomListToString(List<ComparisonUnitAtom> comparisonUnitAtomList, int indent)
        {
            StringBuilder sb = new StringBuilder();
            var cal = comparisonUnitAtomList
                .Select((ca, i) => new
                {
                    ComparisonUnitAtom = ca,
                    Index = i,
                });
            foreach (var item in cal)
                sb.Append("".PadRight(indent))
                  .AppendFormat("[{0:000000}] ", item.Index + 1)
                  .Append(item.ComparisonUnitAtom.ToString(0) + Environment.NewLine);
            return sb.ToString();
        }
    }

    internal enum ComparisonUnitGroupType
    {
        Paragraph,
        Table,
        Row,
        Cell,
    };

    internal class ComparisonUnitGroup : ComparisonUnit
    {
        public ComparisonUnitGroupType ComparisonUnitGroupType;

        public ComparisonUnitGroup(IEnumerable<ComparisonUnit> comparisonUnitList, ComparisonUnitGroupType groupType)
        {
            Contents = comparisonUnitList.ToList();
            ComparisonUnitGroupType = groupType;
            var first = comparisonUnitList.First();
            ComparisonUnitAtom comparisonUnitAtom = GetFirstComparisonUnitAtomOfGroup(first);
            XName ancestorName = null;
            if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Table)
                ancestorName = W.tbl;
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Row)
                ancestorName = W.tr;
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Cell)
                ancestorName = W.tc;
            else if (groupType == OpenXmlPowerTools.ComparisonUnitGroupType.Paragraph)
                ancestorName = W.p;
            var ancestor = comparisonUnitAtom.AncestorElements.Reverse().FirstOrDefault(a => a.Name == ancestorName);
            if (ancestor == null)
                throw new OpenXmlPowerToolsException("Internal error: ComparisonUnitGroup");
            SHA1Hash = (string)ancestor.Attribute(PtOpenXml.SHA1Hash);
        }

        public static ComparisonUnitAtom GetFirstComparisonUnitAtomOfGroup(ComparisonUnit group)
        {
            var thisGroup = group;
            while (true)
            {
                var tg = thisGroup as ComparisonUnitGroup;
                if (tg != null)
                {
                    thisGroup = tg.Contents.First();
                    continue;
                }
                var tw = thisGroup as ComparisonUnitWord;
                if (tw == null)
                    throw new OpenXmlPowerToolsException("Internal error: GetFirstComparisonUnitAtomOfGroup");
                var ca = (ComparisonUnitAtom)tw.Contents.First();
                return ca;
            }
        }

        public override string ToString(int indent)
        {
            var sb = new StringBuilder();
            sb.Append("".PadRight(indent) + "Group Type: " + ComparisonUnitGroupType.ToString() + " SHA1:" + SHA1Hash + Environment.NewLine);
            foreach (var comparisonUnitAtom in Contents)
                sb.Append(comparisonUnitAtom.ToString(indent + 2));
            return sb.ToString();
        }
    }

    public enum CorrelationStatus
    {
        Nil,
        Normal,
        Unknown,
        Inserted,
        Deleted,
        Equal,
        Group,
    }

    class PartSHA1HashAnnotation
    {
        public string Hash;

        public PartSHA1HashAnnotation(string hash)
        {
            Hash = hash;
        }
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
            if (CorrelationStatus == OpenXmlPowerTools.CorrelationStatus.Equal)
            {
                sb.Append(indentString4 + "ComparisonUnitList =====" + Environment.NewLine);
                foreach (var item in ComparisonUnitArray2)
                    sb.Append(item.ToString(6) + Environment.NewLine);
            }
            else
            {
                if (ComparisonUnitArray1 != null)
                {
                    sb.Append(indentString4 + "ComparisonUnitList1 =====" + Environment.NewLine);
                    foreach (var item in ComparisonUnitArray1)
                        sb.Append(item.ToString(6) + Environment.NewLine);
                }
                if (ComparisonUnitArray2 != null)
                {
                    sb.Append(indentString4 + "ComparisonUnitList2 =====" + Environment.NewLine);
                    foreach (var item in ComparisonUnitArray2)
                        sb.Append(item.ToString(6) + Environment.NewLine);
                }
            }
            return sb.ToString();
        }
    }
}
