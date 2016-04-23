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

#define COPY_FILES_FOR_DEBUGGING

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class RsTests
    {
        [Theory]
        [InlineData("HC009-Test-04.docx")]
        public void RS001_Annotations(string name)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

#if COPY_FILES_FOR_DEBUGGING
            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-1-Source.docx")));
            File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var annotatedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-2-Annotated.docx")));
            File.Copy(sourceDocx.FullName, annotatedDocx.FullName);

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(annotatedDocx.FullName, true))
            {
                WmlRunSplitter.Split(wDoc, new[] { wDoc.MainDocumentPart });
            }
            //var assembledFormattingDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-2-FormattingAssembled.docx")));
            //CopyFormattingAssembledDocx(sourceDocx, assembledFormattingDestDocx);
#endif

            //var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-3-OxPt.html")));
            //ConvertToHtml(sourceDocx, oxPtConvertedDestHtml);
        }

        [Theory]
        [InlineData("SR001-Plain.docx")]
        [InlineData("SR002-Bookmark.docx")]
        [InlineData("SR003-Numbered-List.docx")]
        [InlineData("SR004-TwoParas.docx")]
        [InlineData("SR005-Table.docx")]
        [InlineData("SR006-ContentControl.docx")]
        
        public void RS002_ContentAtoms(string name)
        {
            FileInfo sourceDocx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-1-Source.docx")));
            File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var coalescedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-2-Coalesced.docx")));
            File.Copy(sourceDocx.FullName, coalescedDocx.FullName);

            var contentAtomDataFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-3-ContentAtomData.txt")));

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(coalescedDocx.FullName, true))
            {
                WmlRunSplitter.Split(wDoc, new[] { wDoc.MainDocumentPart });
                StringBuilder sb = new StringBuilder();
                foreach (var part in wDoc.ContentParts())
                {
                    var spa = part.Annotation<ContentAtomListAnnotation>();
                    if (spa == null)
                        throw new OpenXmlPowerToolsException("Internal error, annotation does not exist");

                    sb.AppendFormat("Part: {0}", part.Uri.ToString());
                    sb.Append(Environment.NewLine);
                    sb.Append(spa.DumpContentAtomListAnnotation(2));
                    sb.Append(Environment.NewLine);

                    XDocument newMainXDoc = WmlRunSplitter.Coalesce(spa);
                    var partXDoc = wDoc.MainDocumentPart.GetXDocument();
                    partXDoc.Root.ReplaceWith(newMainXDoc.Root);
                }
                File.WriteAllText(contentAtomDataFi.FullName, sb.ToString());
            }
        }
    }
}