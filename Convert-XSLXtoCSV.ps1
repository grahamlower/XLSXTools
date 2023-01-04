param (
    [Parameter(Mandatory=$true, Position = 1)]
    [ValidatePattern('.xlsx$')]
    [System.IO.FileInfo]$InputFile,

    # [Parameter(Mandatory=$false, Position = 2)]
    # [System.IO.FileInfo]$SaveAs,

    #[switch]$overwrite, 

    
    [Parameter(ParameterSetName='ByNum')]
    [int]$Tab = 1,
    
    
    [Parameter(Mandatory = $true, ParameterSetName='ByName')]
    [string]$TabName

    
)

function Get-ArchiveEntryContent {
    param (
        [Parameter(Mandatory=$true)]
        [System.IO.Compression.ZipArchiveEntry]$entry
    )

    try{
        $stream = $entry.Open()
        $reader = [System.IO.StreamReader]::new($stream)

        return $reader.ReadToEnd();
    }
    catch{throw}
    finally{
        if ($null -ne $reader) { $reader.Dispose() }
        $stream.Dispose()
    }

}

function ConvertFrom-R1C1 {
    param (
        [Parameter(Mandatory=$true)]
        [string]$r1c1
    )
                 
    if ($r1c1 -match "((?:[A-Z])+)([1-9][0-9]{0,})") {
        $rows = [int]$matches[2]
        
        $colsABC = $matches[1].ToUpper().ToCharArray()
        
        $cols = 0
        for($c = 0; $c -lt $colsABC.Count; $c++) {
            $power = ($colsABC.Count - $c ) - 1

            $cols += ([int]$colsABC[$c] - 64) * [Math]::Pow(26, $power)
        }

        return [PSCustomObject]@{row = [int]$rows; col = [int]$cols}
    } else {
        throw "Invalid R1C1 reference: $R1C1"
    }


}

function ConvertTo-R1C1 {
    param (
        [Parameter(Mandatory=$true)]
        [int]$row,

        [Parameter(Mandatory=$true)]
        [int]$col
    )

    $C1 = ''

    $d = $col

    while ($d -gt 0) {
    $r = $d % 26
    $d -= $r 
    $d /= 26

    $C1= "$([char]($r + 64))$C1"
    } 
                
    return "$c1$row"
}


# $VerbosePreference = 'SilentlyContinue'
# $DebugPreference = 'SilentlyContinue' 

# if ($null -eq $SaveAs) {
#     Write-Debug $InputFile.FullName 
#     Write-Debug ($InputFile.FullName -Replace '\.xlsx$','.csv')
#     $SaveAs = [System.IO.FileInfo]($InputFile.FullName -Replace '\.xlsx$','.csv')
# }

# if (-not $SaveAs.Directory.Exists) {
#     Throw "Can not file as $SaveAs, $($SaveAs.DirectoryName) does not exist"
# }

# if ($SaveAs.Exists -and -not $overwrite) {
#     Throw "Can not file as $SaveAs, file already exists. Use -OverWrite"
# }

#Default Formats
$numFmt = @{
    0  = @{ formatCode = $null}
    1  = @{ formatCode = '0' }
    2  = @{ formatCode = '0.00' } 
    3  = @{ formatCode = '#,##' } 
    4  = @{ formatCode = '#,##0.00'}
    9  = @{ formatCode = '0%'}
    10 = @{ formatCode = '0.00%'}
    11 = @{ formatCode = '0.00E+00'}
    12 = @{ formatCode = '# ?/?'} #?
    13 = @{ formatCode = '# ??/??'} #?
    14 = @{ formatCode = 'd'; Type = 'DateTime' }
    15 = @{ formatCode = 'd-MMM-yy'; Type = 'DateTime' }
    16 = @{ formatCode = 'd-MMM'; Type = 'DateTime' }
    17 = @{ formatCode = 'MMM-yy'; Type = 'DateTime' }
    18 = @{ formatCode = 't'; Type = 'DateTime' }
    19 = @{ formatCode = 'T'; Type = 'DateTime' }
    20 = @{ formatCode = 'HH:mm'; Type = 'DateTime' }
    21 = @{ formatCode = 'HH:mm:ss'; Type = 'DateTime' }
    22 = @{ formatCode = 'g'; Type = 'DateTime' }
    37 = @{ formatCode = '#,##0 ;(#,##0)'}
    38 = @{ formatCode = '#,##0 ;[Red](#,##0)'}
    39 = @{ formatCode = '#,##0.00;(#,##0.00)'}
    40 = @{ formatCode = '#,##0.00;[Red](#,##0.00)'}
    45 = @{ formatCode = 'mm:ss'; Type = 'DateTime' } 
    46 = @{ formatCode = '[HH]:mm:ss'; Type = 'DateTime' } 
    47 = @{ formatCode = 'mm:ss.0'; Type = 'DateTime' } 
    48 = @{ formatCode = '##0.0E+0'}
    49 = @{ formatCode = '@'} #?
}



try {
    #$zipToOpen = [System.IO.FileStream]::new($InputFile, [System.IO.FileMode]::Open)
    $archive  = [System.IO.Compression.ZipFile]::OpenRead($InputFile)
    
    $contents =  @{}
    $archive.Entries| Where-Object { $_.FullName -match '\.xml(\.rels)?$' } | ForEach-Object { 
        Write-Debug "Loading $($_.FullName)"
        $contents += @{ "$($_.FullName)" = [xml](Get-ArchiveEntryContent $_)  }   
    }
    
    #Add Formats Defined in Document
    $contents['xl/styles.xml'].SelectNodes("//*[local-name() = 'numFmt']") |% { 
        $formatInfo = @{ formatCode = $_.formatCode}
    
        if ($formatInfo.formatCode -match "[mdyhms]" -and (-not $formatInfo.formatCode.Contains(';'))) { $formatInfo.Type = 'DateTime'  }
        
        $numFmt += @{ [int]($_.numFmtId) =$formatInfo } 
        
    }

    $styles = [System.Collections.Generic.List[Object]]::New()
    $styles.AddRange(($contents['xl/styles.xml'].SelectNodes("//*[local-name() = 'cellXfs']").xf|Select-Object numFmtId, applyNumberFormat, quotePrefix))
    
    $styles | Write-Debug 

    #Load the Share Strings
    $sharedStrings = @()
    $sharedStrings += @($contents['xl/sharedStrings.xml'].SelectNodes("//*[local-name() = 'si']").t)
    
    # Find wroksheet tab to export
    if($PSBoundParameters.ContainsKey("TabName")) {
        write-debug "Tab Name: $tabName"
        $worksheetNode = $contents['xl/workbook.xml'].SelectNodes("(//*[local-name()='sheet'])")| ? { $_.Name -like $TabName } | SELECT -First 1

    } else {
        write-debug "(//*[local-name='sheet'])[$Tab]"
        
        #$contents['xl/workbook.xml'].SelectSingleNode("(//*[local-name()='workbook'])[$Tab]").GetType()

        $worksheetNode = $contents['xl/workbook.xml'].SelectSingleNode("(//*[local-name()='sheet'])[$Tab]")
    }

    if ($null -eq $worksheetNode) {
        Throw "Could not find worksheet"
    } else {
        $id = ([xml]($worksheetNode.OuterXml)).sheet.id
    }

    $worksheetPath = ([xml]($contents['xl/_rels/workbook.xml.rels'].SelectSingleNode("//*[@Id='$id']").OuterXml)).Relationship.Target

    Write-Debug "Worksheet Path: $worksheetPath"

    #Load the sheet data into cells, with there "location", for further processing
    $sheetData = ([xml]($contents["xl/$worksheetPath"].worksheet.sheetData.OuterXml)).sheetData

    $cells = [System.Collections.Generic.List[Object]]::New()  # [Script.SortedCells]::New()

    $maxRow= 0
    $maxCol = 0

    $cells.AddRange(($sheetdata.SelectNodes("//*[local-name() = 'v']/..") | Select-Object r,t,s,v,@{name="pos"; Express ={.\ConvertFrom-R1C1.ps1 $_.r}}))

    $size = $cells | Measure-Object -Maximum { $_.pos.row},{$_.pos.col} | SELECT Maximum

    #store the data we are going to output.  Size based upon the data size we have
    $data = New-Object string[][] $size[0].Maximum,$size[1].Maximum  #$size.row,$size.col

    foreach($cell in $cells)  {
        write-debug "--------------"
        write-debug $cell

        $cellType = $cell.t
        $cellValue = $cell.v 
        $cellStyle = [Nullable[int]]$cell.s
 
        <# if the cell has a style (but type if otherwise not indicated) , it is either a date or a number.  If the FormatInfo has a Type of DateTime, it is a date   #>
        if ($null -ne $cellStyle -and $null -eq  $cellType -and $null -ne $styles[$cellStyle] ) {
            Write-Verbose "Applying format information" 

            $formatInfo = $numFmt[[int]($styles[$cellStyle].numFmtId)]
            if($formatInfo.Type -eq 'DateTime') {  $cellType = 'd' } else {  $cellType = 'n'}

        } elseif ($null -eq $cellStyle -and $null -eq  $cellType -and [Double]::TryParse($cellValue,[ref]$null)) {
            $cellType = 'n'
            $formatInfo = ${formatCode = 'G11'}
        } else {
            $formatInfo = ${formatCode = $null}
        }
        
        
        #Formatting data 
        $v = switch($cellType) {
            'b' { if($cellValue -like 'y*' -or $cellValue -like 't*') { "TRUE" } else { 'FALSE'} <# boolean #>} 
            's' { "`"$($sharedStrings[([int]$cellValue)])`""  <#sharedStrings#> }
            'd' {  
                $numericValue = [float]$cellValue
                $formatCode = $formatInfo.formatCode
                Write-Debug "numericValue: $numericValue"
                Write-Debug "formatCode: $formatCode"

                # ECMA-376, 4th Edition, Part 1, 18.8.31 Special language info values
                if ($formatCode.StartsWith('[$-F400]')) {$formatCode = 'T' }
                elseif ($formatCode.StartsWith('$-F800]*')) {$formatCode = 'D' }

                if ($numericValue -eq 60) {
                    (get-date '1/1/1900').AddDays(-1).AddDays(59).ToString($formatCode).Replace('28', '29')
                } else {
                    if ($numericValue -gt 60) {$numericValue-- <# 02/29/1900 fix #>}
                    $baseDate = (get-date '1/1/1900').AddDays(-1).AddDays($numericValue)
                    Write-Debug "Date: $baseDate"
                    $baseDate.ToString($formatCode)
                }
            }
            'e' { "error"}
            'inlineStr' { 'rich string'}
            'n' {  
                $numericValue = [double]$cellValue
                Write-Debug "numericValue: $numericValue"
                Write-Debug "formatCode: $($formatInfo.formatCode)"
                $numericValue -f $formatInfo.formatCode

            }
            # 'str' { "formula string"
        
            # }
            Default { "`"$cellValue`"" }
        }

        write-debug "V: $v"



        $data[$cell.pos.row - 1][$cell.pos.col - 1] = $v
    }

    $data | % { $_ -join ',' }
    

}
catch {
    Throw
}
try {
    
}
finally {
   if($null -ne $archive) { $archive.Dispose()}
   if($null -ne $zipToOpen) { $zipToOpen.Dispose() }
}
