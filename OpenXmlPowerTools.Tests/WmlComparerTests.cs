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
using OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class WcTests
    {
        [Theory]
        [InlineData("SR001-Plain.docx", "SR001-Plain-Mod.docx")]
        [InlineData("WC001-Digits.docx", "WC001-Digits-Mod.docx")]

        public void WC001_Compare(string name1, string name2)
        {
            FileInfo source1Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name1));
            FileInfo source2Docx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name2));

            var source1CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source1Docx.Name.Replace(".docx", "-1-Source.docx")));
            var source2CopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source2Docx.Name.Replace(".docx", "-2-Source.docx")));
            File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
            File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

            var docxWithRevisionsFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, source1Docx.Name.Replace(".docx", "-3-Compared.docx")));

            WmlDocument source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
            WmlDocument source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
            WmlComparerSettings settings = new WmlComparerSettings();
            WmlDocument comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);
            comparedWml.SaveAs(docxWithRevisionsFi.FullName);
        }
    }
}