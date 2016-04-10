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
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;

namespace OpenXmlPowerTools
{
    public class WmlComparerSettings
    {
        public char[] WordSeparators;
        public bool AcceptRevisionsBeforeProcessing;

        public WmlComparerSettings()
        {
            // note that , and . are processed explicitly to handle cases where they are in a number or word
            WordSeparators = new[] { ' ' }; // todo need to fix this for complete list
            AcceptRevisionsBeforeProcessing = true;
        }
    }

    public static class WmlComparer
    {
        // todo need to accept revisions if necessary
        // todo look for invalid content, throw if found
        public static WmlDocument Compare(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings)
        {
            WmlDocument wmlResult = new WmlDocument(source1);
            using (MemoryStream ms1 = new MemoryStream())
            using (MemoryStream ms2 = new MemoryStream())
            {
                ms1.Write(source1.DocumentByteArray, 0, source1.DocumentByteArray.Length);
                ms2.Write(source2.DocumentByteArray, 0, source2.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc1 = WordprocessingDocument.Open(ms1, true))
                using (WordprocessingDocument wDoc2 = WordprocessingDocument.Open(ms2, true))
                {
                    WmlRunSplitter.Split(wDoc1);
                    WmlRunSplitter.Split(wDoc2);
                    SplitRunsAnnotation sra1 = wDoc1.MainDocumentPart.Annotation<SplitRunsAnnotation>();
                    SplitRunsAnnotation sra2 = wDoc2.MainDocumentPart.Annotation<SplitRunsAnnotation>();
                    return ApplyChanges(sra1, sra2, wmlResult, settings);
                }
            }
        }

        private static WmlDocument ApplyChanges(SplitRunsAnnotation sra1, SplitRunsAnnotation sra2, WmlDocument wmlResult,
            WmlComparerSettings settings)
        {
            var cu1 = GetComparisonUnitList(sra1, settings);
            var cu2 = GetComparisonUnitList(sra2, settings);
            var correlatedSequence = Lcs(cu1, cu2);

            var sb = new StringBuilder();
            foreach (var item in correlatedSequence)
                sb.Append(item.ToString());
            var s = sb.ToString();
            Console.WriteLine(s);

            return null;
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

            var sb = new StringBuilder();
            foreach (var item in csList)
                sb.Append(item.ToString());
            var s = sb.ToString();
            Console.WriteLine(s);


            while (true)
            {
                var unknown = csList
                    .FirstOrDefault(z => z.CorrelationStatus == CorrelationStatus.Unknown);
                if (unknown == null)
                    break;



                // do LCS on paragraphs here




                List<CorrelatedSequence> newSequence = FindLongestCommonSequence(unknown);
                var indexOfUnknown = csList.IndexOf(unknown);
                csList.Remove(unknown);

                newSequence.Reverse();
                foreach (var item in newSequence)
                    csList.Insert(indexOfUnknown, item);
            }

            return csList;
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
            }
            else if (currentI1 == 0 && currentI2 > 0)
            {
                var insertedCorrelatedSequence = new CorrelatedSequence();
                insertedCorrelatedSequence.CorrelationStatus = CorrelationStatus.Inserted;
                insertedCorrelatedSequence.ComparisonUnitArray1 = null;
                insertedCorrelatedSequence.ComparisonUnitArray2 = cul2.Take(currentI2).ToArray();
                newListOfCorrelatedSequence.Add(insertedCorrelatedSequence);
            }
            else if (currentI1 > 0 && currentI2 > 0)
            {
                var unknownCorrelatedSequence = new CorrelatedSequence();
                unknownCorrelatedSequence.CorrelationStatus = CorrelationStatus.Unknown;
                unknownCorrelatedSequence.ComparisonUnitArray1 = cul1.Take(currentI1).ToArray();
                unknownCorrelatedSequence.ComparisonUnitArray2 = cul2.Take(currentI2).ToArray();
                newListOfCorrelatedSequence.Add(unknownCorrelatedSequence);
            }
            else if (currentI1 == 0 && currentI2 == 0)
            {
                var equalCorrelatedSequence = new CorrelatedSequence();
                equalCorrelatedSequence.CorrelationStatus = CorrelationStatus.Equal;
                equalCorrelatedSequence.ComparisonUnitArray1 = cul1.Skip(currentI1).Take(currentLongestCommonSequenceLength).ToArray();
                equalCorrelatedSequence.ComparisonUnitArray2 = cul2.Skip(currentI2).Take(currentLongestCommonSequenceLength).ToArray();
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
                var equalCorrelatedSequence = new CorrelatedSequence();
                equalCorrelatedSequence.CorrelationStatus = CorrelationStatus.Equal;
                equalCorrelatedSequence.ComparisonUnitArray1 = cul1.Skip(currentI1).ToArray();
                equalCorrelatedSequence.ComparisonUnitArray2 = cul2.Skip(currentI2).ToArray();
                newListOfCorrelatedSequence.Add(equalCorrelatedSequence);
            }
            return newListOfCorrelatedSequence;
        }

        private static ComparisonUnit[] GetComparisonUnitList(SplitRunsAnnotation splitRunsAnnotation, WmlComparerSettings settings)
        {
            var splitRuns = splitRunsAnnotation.SplitRuns;
            var groupingKey = splitRuns
                .Select((sr, i) =>
                {
                    string key = null;
                    if (sr.ContentAtom.Name == W.t)
                    {
                        string chr = sr.ContentAtom.Value;
                        var ch = chr[0];
                        if (ch == '.' || ch == ',')
                        {
                            bool beforeIsDigit = false;
                            if (i > 0)
                            {
                                var prev = splitRuns[i - 1];
                                if (prev.ContentAtom.Name == W.t && char.IsDigit(prev.ContentAtom.Value[0]))
                                    beforeIsDigit = true;
                            }
                            bool afterIsDigit = false;
                            if (i < splitRuns.Length - 1)
                            {
                                var next = splitRuns[i + 1];
                                if (next.ContentAtom.Name == W.t && char.IsDigit(next.ContentAtom.Value[0]))
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
                    else if (sr.ContentAtom.Name == W.pPr)
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
                        SplitRunMember = sr
                    };
                });

            //var sb = new StringBuilder();
            //foreach (var item in groupingKey)
            //{
            //    sb.Append(item.Key + Environment.NewLine);
            //    sb.Append("    " + item.SplitRunMember.ToString(0) + Environment.NewLine);
            //}
            //var sbs = sb.ToString();
            //Console.WriteLine(sbs);

            var groupedByWords = groupingKey
                .GroupAdjacent(gc => gc.Key);

            //var sb = new StringBuilder();
            //foreach (var group in groupedByWords)
            //{
            //    sb.Append("Group ===== " + group.Key + Environment.NewLine);
            //    foreach (var gc in group)
            //    {
            //        sb.Append("    " + gc.SplitRunMember.ToString(0) + Environment.NewLine);
            //    }
            //}
            //var sbs = sb.ToString();
            //Console.WriteLine(sbs);

            ComparisonUnit[] cul = groupedByWords
                .Select(g => new ComparisonUnit(g.Select(gc => gc.SplitRunMember)))
                .ToArray();

            //var sb = new StringBuilder();
            //foreach (var group in cul)
            //{
            //    sb.Append("Group ===== " + Environment.NewLine);
            //    foreach (var gc in group.Contents)
            //    {
            //        sb.Append("    " + gc.ToString(0) + Environment.NewLine);
            //    }
            //}
            //var sbs = sb.ToString();
            //Console.WriteLine(sbs);

            return cul;
        }
    }

    internal class ComparisonUnit : IEquatable<ComparisonUnit>
    {
        public List<SplitRun> Contents;
        public ComparisonUnit(IEnumerable<SplitRun> splitRuns)
        {
            Contents = splitRuns.ToList();
        }

        public string ToString(int indent)
        {
            var sb = new StringBuilder();
            foreach (var splitRun in Contents)
                sb.Append(splitRun.ToString(indent) + Environment.NewLine);
            return sb.ToString();
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

            if (this.Contents.Any(c => c.ContentAtom.Name == W.t) ||
                other.Contents.Any(c => c.ContentAtom.Name == W.t))
            {
                var txt1 = this
                    .Contents
                    .Where(c => c.ContentAtom.Name == W.t)
                    .Select(c => c.ContentAtom.Value)
                    .StringConcatenate();
                var txt2 = other
                    .Contents
                    .Where(c => c.ContentAtom.Name == W.t)
                    .Select(c => c.ContentAtom.Value)
                    .StringConcatenate();
                if (txt1 != txt2)
                    return false;

                var seq1 = this
                    .Contents
                    .Where(c => ! s_ElementsToIgnoreWhenComparing.Contains(c.ContentAtom.Name));
                var seq2 = this
                    .Contents
                    .Where(c => ! s_ElementsToIgnoreWhenComparing.Contains(c.ContentAtom.Name));
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
                    .Where(c => !s_ElementsToIgnoreWhenComparing.Contains(c.ContentAtom.Name));
                var seq2 = this
                    .Contents
                    .Where(c => !s_ElementsToIgnoreWhenComparing.Contains(c.ContentAtom.Name));
                if (seq1.Count() != seq2.Count())
                    return false;
                var zipped = seq1.Zip(seq2, (s1, s2) => new
                {
                    Cu1 = s1,
                    Cu2 = s2,
                });
                var anyNotEqual = (zipped.Any(z =>
                {
                    if (z.Cu1.ContentAtom.Name != z.Cu2.ContentAtom.Name)
                        return false;
                    var a1 = z.Cu1.AncestorElements.Select(a => a.Name.ToString() + "|").StringConcatenate();
                    var a2 = z.Cu2.AncestorElements.Select(a => a.Name.ToString() + "|").StringConcatenate();
                    return a1 != a2;
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
