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

            return null;
        }

        private static List<ComparisonUnit> GetComparisonUnitList(SplitRunsAnnotation splitRunsAnnotation, WmlComparerSettings settings)
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

            List<ComparisonUnit> cul = groupedByWords
                .Select(g => new ComparisonUnit(g.Select(gc => gc.SplitRunMember)))
                .ToList();

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

    // tomorrow first thing - write xunits for ComparisonUnit equals

    internal class ComparisonUnit : IEquatable<ComparisonUnit>
    {
        public List<SplitRun> Contents;
        public ComparisonUnit(IEnumerable<SplitRun> splitRuns)
        {
            Contents = splitRuns.ToList();
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
}
