<#***************************************************************************

Copyright (c) Microsoft Corporation 2014.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

Version: 3.0.0

***************************************************************************#>

function Convert-HtmlToDocx {
    <#
    .SYNOPSIS
     Converts a HTML/CSS to DOCX.
    .DESCRIPTION
     Converts a HTML/CSS to DOCX outputting images to a related directory.
    .PARAMETER FileName
     The HTML file to convert to DOCX.
    .PARAMETER OutputPath
     The directory that will contain the converted DOCX file.
    .PARAMETER EmptyWordDocument
     Creates empty word document. The HTML contents will be saved in this document
    .PARAMETER CustomCss
     Use this parameter to specify the user defined CSS/custom Css for the resultant document
    .PARAMETER OpenNow
     Use this parameter to open the document immediately.
    .PARAMETER OpenNow
     Use this parameter to create annotation file
    .EXAMPLE
     convert-htmltodocx Resume.html -OutputPath demo.docx
    .EXAMPLE
     convert-htmltodocx sample.html -OutputPath demo.docx -OpenNow -EmptyWordDocument
    .EXAMPLE
     Convert-HtmlToDocx *.html -OutputPath .\test\ -EmptyWordDocument -OpenNow
    .EXAMPLE
     convert-htmltodocx sample.html -OutputPath demo.docx -EmptyWordDocument
    .EXAMPLE
     convert-htmltodocx sample.html -OutputPath demo.docx -OpenNow
    .EXAMPLE
     Convert-HtmlToDocx Resume.html -EmptyWordDocument -OpenNow -CreateAnnotation
    .EXAMPLE
     Convert-htmltodocx sample.html -outputpath sample.docx
    .EXAMPLE
     Convert-htmltodocx sample.html -outputpath .\test\sample.docx
    .EXAMPLE
     Convert-HtmlToDocx sample.html -EmptyWordDocument -OpenNow
    .EXAMPLE
     Convert-HtmlToDocx sample.html -OutputPath demo.docx -CustomCss C:\custom.css
    .EXAMPLE
     Convert-HtmlToDocx sample.html -CustomCss C:\custom.css -OpenNow -EmptyWordDocument
    .EXAMPLE
     convert-htmltodocx Resume.html -OutputPath demo.docx -OpenNow -CreateAnnotation
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Which HTML file would you like to transform to DOCX?',
        Position=0)]
        [ValidateScript(
        {
            $prevCurrentDirectory = [Environment]::CurrentDirectory
            [environment]::CurrentDirectory = $(Get-Location)
            if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_))
            {
            	[environment]::CurrentDirectory = $prevCurrentDirectory
                return $True
            }
            else
            {
                if (Test-Path $_)
                {
                	[environment]::CurrentDirectory = $prevCurrentDirectory
                    return $True
                }
                else
                {
                	[environment]::CurrentDirectory = $prevCurrentDirectory
                    Throw "$_ is not a valid filename"
                }
            }
        })]
        [SupportsWildcards()]
        [string[]]$FileName,
        
        [Parameter(Mandatory=$False,
        Position=1,
        ValueFromPipeline=$False)]
        [ValidateScript(
        {
            $prevCurrentDirectory = [Environment]::CurrentDirectory
            [environment]::CurrentDirectory = $(Get-Location)

            if (Test-Path $_)
            {
                [environment]::CurrentDirectory = $prevCurrentDirectory
                Throw "$_ already exists"
            }
            else
            {
                [environment]::CurrentDirectory = $prevCurrentDirectory
                return $True
            }
        })]
        [string]$OutputPath,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$EmptyWordDocument,

        [Parameter(Mandatory=$False)]
        [string]$CustomCss,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$OpenNow,

        [Parameter(Mandatory=$False)]
        [Switch]
        [bool]$CreateAnnotation

    )

   
    begin {
        $prevCurrentDirectory = [Environment]::CurrentDirectory
        [environment]::CurrentDirectory = $(Get-Location)
    }
   
    process {
        if ($OutputPath -ne [string]::Empty)
        {
            foreach ($argItem in $FileName) {
                if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($argItem))
                {
                    $dir = New-Object -TypeName System.IO.DirectoryInfo $(Get-Location)
                    foreach ($fi in $dir.GetFiles($argItem))
                    {
                        Convert-HtmlTODocx-Helper $fi $OutputPath $EmptyWordDocument $CustomCss $OpenNow $CreateAnnotation
                    }
                }
                else
                {
                    $fi = New-Object System.IO.FileInfo($argItem)
                    Convert-HtmlTODocx-Helper $fi $OutputPath $EmptyWordDocument $CustomCss $OpenNow $CreateAnnotation
                }
            }
        }
        else
        {
            if ($EmptyWordDocument -eq $True)
            {
                [OpenXmlPowerTools.HtmlToDocxConverterHelper]::ConvertHtmlToDocx($FileName,$OutputPath,$EmptyWordDocument, $CustomCss, $OpenNow, $CreateAnnotation)
            }
            else
            {
                Throw "Output path does not exists!"
            }
        }
    }

    end {
        [environment]::CurrentDirectory = $prevCurrentDirectory
    }
}

function Convert-HtmlTODocx-Helper {
    param (
        [System.IO.FileInfo]$fi,
        [string]$OutputPath,
        [bool]$EmptyWordDocument,
        [string]$CustomCss,
        [bool]$OpenNow,
        [bool]$CreateAnnotation
    )
   [OpenXmlPowerTools.HtmlToDocxConverterHelper]::ConvertHtmlToDocx($fi.FullName,$OutputPath,$EmptyWordDocument, $CustomCss, $OpenNow,$CreateAnnotation);
}
