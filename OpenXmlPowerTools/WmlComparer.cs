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
using System.Drawing;

namespace OpenXmlPowerTools
{
    public class WmlComparerSettings
    {
        public char[] WordSeparators;

        public WmlComparerSettings()
        {
            WordSeparators = new char[] { };
        }
    }

    public static class WmlComparer
    {
        public static WmlDocument Compare(WmlDocument source1, WmlDocument source2, WmlComparerSettings settings)
        {
            //using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(document))
            //{
            //    using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
            //    {
            //        AssembleFormatting(doc, settings);
            //    }
            //    return streamDoc.GetModifiedWmlDocument();
            //}
            return null;
        }

        private static WmlDocument Compare(WordprocessingDocument source1, WordprocessingDocument source2, WmlComparerSettings settings)
        {
            return null;
        }
    }
}
