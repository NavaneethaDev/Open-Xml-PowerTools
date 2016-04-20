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
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class WcTests
    {
        public static string[] ExpectedErrors = new string[] {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:latentStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:stylesInUse' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:headingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:numberingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tableStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnRuns' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnParagraphs' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnNumbering' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnTables' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:clearFormatting' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:top3HeadingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:visibleStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:alternateStyleNames' attribute is not declared.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 1.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 2.",
        };

        [Theory]
        [InlineData("SR001-Plain.docx", "SR001-Plain-Mod.docx")]
        [InlineData("WC001-Digits.docx", "WC001-Digits-Mod.docx")]
        [InlineData("WC001-Digits.docx", "WC001-Digits-Deleted-Paragraph.docx")]
        [InlineData("WC001-Digits-Deleted-Paragraph.docx", "WC001-Digits.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DiffInMiddle.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DiffAtBeginning.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DeleteAtBeginning.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-InsertAtBeginning.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-InsertAtEnd.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DeleteAtEnd.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-DeleteInMiddle.docx")]
        [InlineData("WC002-Unmodified.docx", "WC002-InsertInMiddle.docx")]
        [InlineData("WC002-DeleteInMiddle.docx", "WC002-Unmodified.docx")]
        [InlineData("WC003-ITU-Document-Delete-Me.docx", "WC003-ITU-Document-Delete-Me-Mod.docx")]
        [InlineData("WC004-Large.docx", "WC004-Large-Mod.docx")]
        [InlineData("WC003-ITU-Document-Delete-Me.docx", "WC003-ITU-Document-Delete-Me-Mod2.docx")]
        [InlineData("WC005-ITU-Document-Delete-Me-Small.docx", "WC005-ITU-Document-Delete-Me-Small-Mod.docx")]
        [InlineData("WC006-Table.docx", "WC006-Table-Delete-Row.docx")]
        [InlineData("WC006-Table.docx", "WC006-Table-Delete-Contests-of-Row.docx")]
        [InlineData("WC008-ITU-Document-Delete-Me-Drawing.docx", "WC008-ITU-Document-Delete-Me-Drawing-Mod.docx")]
        [InlineData("WC007-Unmodified.docx", "WC007-Longest-At-End.docx")]
        [InlineData("WC007-Unmodified.docx", "WC007-Deleted-at-Beginning-of-Para.docx")]
        [InlineData("WC007-Unmodified.docx", "WC007-Moved-into-Table.docx")]
        [InlineData("WC009-Table-Unmodified.docx", "WC009-Table-Cell-1-1-Mod.docx")]
        [InlineData("WC010-Para-Before-Table-Unmodified.docx", "WC010-Para-Before-Table-Mod.docx")]
        [InlineData("WC011-Before.docx", "WC011-After.docx")]

        public void WC001_Compare(string name1, string name2)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source1Docx.Name.Replace(".docx", "-1-Source.docx")));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source2Docx.Name.Replace(".docx", "-2-Source.docx")));
            if (!source1CopiedToDestDocx.Exists)
                File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            if (!source2CopiedToDestDocx.Exists)
                File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            var before = source1CopiedToDestDocx.Name.Replace(".docx", "");
            var after = source2CopiedToDestDocx.Name.Replace(".docx", "");
            var docxWithRevisionsFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, before + "-COMPARE-" + after + ".docx"));

            WmlDocument source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            comparedWml.SaveAs(docxWithRevisionsFi.FullName);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedWml.DocumentByteArray, 0, comparedWml.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator validator = new OpenXmlValidator();
                    var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
                    if (errors.Count() == 0)
                        return;

                    var ind = "  ";
                    var sb = new StringBuilder();
                    foreach (var err in errors)
                    {
#if true
                        sb.Append("Error" + Environment.NewLine);
                        sb.Append(ind + "ErrorType: " + err.ErrorType.ToString() + Environment.NewLine);
                        sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                        sb.Append(ind + "Part: " + err.Part.Uri.ToString() + Environment.NewLine);
                        sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
#else
                        sb.Append("            \"" + err.Description + "\"," + Environment.NewLine);
#endif

                    }
                    var sbs = sb.ToString();
                    Assert.Equal("", sbs);
                }
            }
        }
    }
}