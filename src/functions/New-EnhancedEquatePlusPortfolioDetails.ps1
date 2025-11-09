<#
    Copyright 2025 Adam Gabryś

    Licensed under the Apache License, Version 2.0 (the "License");
    you may not use this file except in compliance with the License.
    You may obtain a copy of the License at:

        http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing, software
    distributed under the License is distributed on an "AS IS" BASIS,
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
    See the License for the specific language governing permissions and
    limitations under the License.
#>

<#
.SYNOPSIS
Generates an enhanced Excel report for an EquatePlus portfolio with detailed worksheets and calculated columns.

.DESCRIPTION
The New-EnhancedEquatePlusPortfolioDetails function reads an EquatePlus portfolio Excel file
and produces a new Excel workbook with multiple worksheets:

- Overview: Summarizes the portfolio by aggregating key metrics.
- Tax Rates: Stores predefined tax rates used in calculations for estimating taxes and net profit.
- Detailed Data: Contains processed portfolio items with formulas for computing financial metrics.
- Input Data: Contains the original source data with formatted dates, serving as a reference
  for the processed information.

Fields Legend:

- Contribution Type: The type of contribution, e.g., Locked Award, Granted Award, Own Contribution, or Company Match.
- Cost Basis: The cost basis of the shares, calculated as Shares × Purchase Price (if applicable).
- Date: The effective date of the allocation or award.
- Estimated Net Profit: Net profit after tax, calculated as Real Unrealized Gains − Estimated Tax.
- Estimated Tax: Estimated tax liability on taxable gains based on predefined tax rates.
- Market Price: The current market price of a single share at the time of processing.
- Own Costs: The portion of costs attributable to the employee, including company match adjustments.
- Percentage Gain: Estimated Net Profit expressed as a percentage of Own Costs.
- Plan: The name of the portfolio plan.
- Real Unrealized Gains: Actual unrealized gains after deducting Own Costs.
- Shares: The number of allocated shares.
- Taxable Unrealized Gains: Gains subject to taxation, calculated as Value − Cost Basis.
- Value: The market value of the allocated shares, calculated as Shares × Market Price.

The function ensures the required version of the ImportExcel module (7.8.10) is installed and imported
during execution, restoring any previously loaded versions afterward.

.PARAMETER InputFile
Path to the source Excel file containing portfolio details.
This parameter is mandatory and can accept input from the pipeline.

.PARAMETER OutputFile
Optional path where the enhanced Excel report will be created.
If not specified, the output file will default to "Enhanced-" followed by the input file name.
This parameter cannot be used when the function receives input from the pipeline.

.INPUTS
System.String
- InputFile: Path to an existing Excel file containing portfolio data. Can be provided via the pipeline.
- OutputFile: Optional path where the enhanced report will be saved. Cannot be used with pipeline input.

.OUTPUTS
System.String
- The full path of the created enhanced Excel report file.

.EXAMPLE
New-EnhancedEquatePlusPortfolioDetails -InputFile 'Portfolio.xlsx' -OutputFile 'DetailedPortfolio.xlsx'

Creates a detailed Excel report from Portfolio.xlsx and saves it as DetailedPortfolio.xlsx.

.EXAMPLE
'Portfolio.xlsx' | New-EnhancedEquatePlusPortfolioDetails

Accepts the input file from the pipeline and generates the enhanced report using the default output file name.

.EXAMPLE
New-EnhancedEquatePlusPortfolioDetails -InputFile 'Portfolio.xlsx' -OutputFile 'DetailedPortfolio.xlsx' -Verbose

Shows detailed progress for module loading, row processing, and worksheet creation.

.NOTES
Author:  Adam Gabryś
Date:    2025-11-09
Version: 0.1.0
License: Apache-2.0

.LINK
https://github.com/agabrys/equateplus-portfolio-details-enhancer
#>
function New-EnhancedEquatePlusPortfolioDetails {
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string]$InputFile,
    [Parameter()]
    [string]$OutputFile
  )
  begin {
    function Use-ImportExcelModuleRequiredVersion {
      param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
      )

      $name = 'ImportExcel'
      $version = '7.8.10'

      $module = Get-Module -Name $name -ListAvailable | Where-Object -Property Version -EQ $version
      if ($null -eq $module) {
        Write-Verbose "Required version ${version} of the ImportExcel module is not installed."
        Install-Module -Name $name -RequiredVersion $version -Scope CurrentUser -Force -AllowClobber
      } else {
        Write-Verbose "Required version ${version} of the ImportExcel module is installed."
      }

      $originalModule = Get-Module -Name $name
      if ($null -eq $originalModule) {
        Write-Verbose 'ImportExcel module is not currently imported.'
      } else {
        Write-Verbose "The currently imported version of the ImportExcel module is $($originalModule.Version)."
      }

      $result = $null
      try {
        if ($null -ne $originalModule -and $originalModule.Version -ne $version) {
          Write-Verbose "Removing the currently imported version $($originalModule.Version) of the ImportExcel module as version $version is needed."
          Remove-Module -Name $name
        }
        if ($null -eq $originalModule -or $originalModule.Version -ne $version) {
          Write-Verbose "Importing version ${version} of the ImportExcel module."
          Import-Module -Name $name -RequiredVersion $version
        }

        $result = & $ScriptBlock

      } finally {
        if ($null -eq $originalModule -or $originalModule.Version -ne $version) {
          Write-Verbose "Restoring the ImportExcel module to its state before the scriptlet was executed."
          Write-Verbose "Removing the currently imported version ${version} of the ImportExcel module."
          Remove-Module -Name $name
          if ($null -ne $originalModule) {
            Write-Verbose "Importing version $($originalModule.Version) of the ImportExcel module."
            Import-Module $originalModule.Name -RequiredVersion $originalModule.Version
          }
        }
      }
      return $result
    }

    function ConvertTo-InternalItem {
      param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$ExcelRow,
        [Parameter(Mandatory = $true)]
        [int]$Index
      )

      Write-Verbose "Parsing row ${Index}:"
      Write-Verbose "  Raw data: ${ExcelRow}."

      $item = [PSCustomObject]@{
        Plan             = $ExcelRow.Plan
        ContributionType = $ExcelRow.'Contribution type'
        PurchasePrice    = [decimal]$ExcelRow.'Strike price / Cost basis'
        MarketPrice      = [decimal]$ExcelRow.'Market price'
        Date             = (Get-Date '1899-12-30').AddDays($ExcelRow.'Available from').ToString('yyyy-MM-dd')
        Shares           = [decimal]$ExcelRow.'Allocated quantity'
      }

      if ($item.ContributionType -eq 'Award') {
        if ($item.PurchasePrice -eq 0) {
          $item.ContributionType = 'Locked Award'
        } else {
          $item.ContributionType = 'Granted Award'
        }
      } elseif ($item.ContributionType -eq 'Purchase') {
        $item.ContributionType = 'Own Contribution'
      } elseif ($item.ContributionType -eq 'Company match') {
        $item.ContributionType = 'Company Match'
      } else {
        Write-Error "Invalid contribution type `"$($item.ContributionType)`" in row ${Index}."
        exit 1
      }

      Write-Verbose "  Parsed data: ${item}."
      return $item
    }

    function New-CleanFileStructure {
      param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
      )
      $fullPath = if ([System.IO.Path]::IsPathRooted($FilePath)) { $FilePath } else { Join-Path -Path (Get-Location) -ChildPath $FilePath }
      $parentDir = [System.IO.Path]::GetDirectoryName($fullPath)
      if (-not (Test-Path -Path $parentDir)) {
        New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
      }
      if (Test-Path -Path $fullPath) {
        Remove-Item -Path $fullPath -Force | Out-Null
      }
      return $fullPath
    }

    function ConvertTo-ItemsWithCellReferences {
      param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject]$InputObject,
        [Parameter(Mandatory = $true)]
        [int]$StartIndex
      )
      begin {
        $items = [System.Collections.ArrayList]@()
        $itemIndex = $StartIndex
      }
      process {
        $items.Add($InputObject) | Out-Null

        $properties = $InputObject.PSObject.Properties

        $letterMap = @{}
        $propertyIndex = 0
        $enumerator = $properties.GetEnumerator()
        while ($enumerator.MoveNext()) {
          $property = $enumerator.Current
          $letterMap[$property.Name] = [char](65 + $propertyIndex)
          $propertyIndex++
        }

        $enumerator = $properties.GetEnumerator()
        while ($enumerator.MoveNext()) {
          $property = $enumerator.Current
          if ($property.Value -is [string]) {
            $newValue = $property.Value -replace '{{index}}', $itemIndex
            foreach ($name in $letterMap.Keys) {
              $newValue = $newValue -replace "{{${name}}}", $letterMap[$name]
            }
            $InputObject.$($property.Name) = $newValue
          }
        }

        $itemIndex++
      }
      end {
        return $items
      }
    }

    function ConvertTo-DetailedItems {
      param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSCustomObject]$InputObject
      )
      begin {
        $items = [System.Collections.ArrayList]@()
      }
      process {
        $items.Add([PSCustomObject]@{
            Date                       = $InputObject.Date
            Plan                       = $InputObject.Plan
            'Contribution Type'        = $InputObject.ContributionType
            Shares                     = $InputObject.Shares
            'Market Price'             = $InputObject.MarketPrice
            'Value'                    = '=ROUND({{Shares}}{{index}} * {{Market Price}}{{index}}, 2)'
            'Purchase Price'           = if ($InputObject.PurchasePrice -gt 0) { $InputObject.PurchasePrice } else { '' }
            'Cost Basis'               = '=IF({{Purchase Price}}{{index}} <> "", ROUND({{Shares}}{{index}} * {{Purchase Price}}{{index}}, 2), "")'
            'Own Costs'                = '=IF({{Contribution Type}}{{index}} = "Company Match", ROUND(''Tax Rates''!C2 * {{Cost Basis}}{{index}}, 2), IF({{Contribution Type}}{{index}} = "Locked Award", "", {{Cost Basis}}{{index}}))'
            'Taxable Unrealized Gains' = '=IF({{Cost Basis}}{{index}} = "", "", ROUND({{Value}}{{index}} - {{Cost Basis}}{{index}}, 2))'
            'Real Unrealized Gains'    = '=IF({{Cost Basis}}{{index}} = "", "", ROUND({{Value}}{{index}} - {{Own Costs}}{{index}}, 2))'
            'Estimated Tax'            = '=IF({{Taxable Unrealized Gains}}{{index}} = "", "", IF({{Taxable Unrealized Gains}}{{index}} < 0, 0, ROUND(''Tax Rates''!C3 * {{Taxable Unrealized Gains}}{{index}}, 2))'
            'Estimated Net Profit'     = '=IF({{Cost Basis}}{{index}} = "", "", ROUND({{Real Unrealized Gains}}{{index}} - {{Estimated Tax}}{{index}}, 2))'
          }) | Out-Null
      }
      end {
        return $items | ConvertTo-ItemsWithCellReferences -StartIndex 2
      }
    }

    function New-OverviewItems {
      param(
        [Parameter(Mandatory = $true)]
        [int]$LastRowIndex
      )

      $items = [System.Collections.ArrayList]@()

      foreach ($item in @(
          [PSCustomObject]@{
            Type   = 'Locked Awards'
            Filter = 'Locked Award'
          },
          [PSCustomObject]@{
            Type   = 'Granted Awards'
            Filter = 'Granted Award'
          },
          [PSCustomObject]@{
            Type   = 'Own Contributions'
            Filter = 'Own Contribution'
          },
          [PSCustomObject]@{
            Type   = 'Company Matches'
            Filter = 'Company Match'
          }
        )) {
        $items.Add([PSCustomObject]@{
            'Contribution Type'        = $item.Type
            Shares                     = "=SUMIF('Detailed Data'!C2:C${LastRowIndex}`, `"$($item.Filter)`", 'Detailed Data'!D2:D${LastRowIndex})"
            Value                      = '=ROUND({{Shares}}{{index}} * ''Detailed Data''!E2, 2)'
            'Cost Basis'               = if ($item.Filter -eq 'Locked Award') { '' } else { "=SUMIF('Detailed Data'!C2:C${LastRowIndex}`, `"$($item.Filter)`", 'Detailed Data'!H2:H${LastRowIndex})" }
            'Own Costs'                = '=IF({{Contribution Type}}{{index}} = "Company Matches", ROUND(''Tax Rates''!C2 * {{Cost Basis}}{{index}}, 2), IF({{Contribution Type}}{{index}} = "Locked Awards", "", {{Cost Basis}}{{index}}))'
            'Taxable Unrealized Gains' = '=IF({{Cost Basis}}{{index}} = "", "", ROUND({{Value}}{{index}} - {{Cost Basis}}{{index}}, 2))'
            'Real Unrealized Gains'    = '=IF({{Cost Basis}}{{index}} = "", "", ROUND({{Value}}{{index}} - {{Own Costs}}{{index}}, 2))'
            'Estimated Tax'            = '=IF({{Taxable Unrealized Gains}}{{index}} = "", "", IF({{Taxable Unrealized Gains}}{{index}} < 0, 0, ROUND(''Tax Rates''!C3 * {{Real Unrealized Gains}}{{index}}, 2))'
            'Estimated Net Profit'     = '=IF({{Cost Basis}}{{index}} = "", "", ROUND({{Real Unrealized Gains}}{{index}} - {{Estimated Tax}}{{index}}, 2))'
            'Percentage Gain'          = '=IF({{Cost Basis}}{{index}} = "", "", ROUND({{Estimated Net Profit}}{{index}} / {{Own Costs}}{{index}} * 100, 2))'
          }) | Out-Null
      }

      $items.Add([PSCustomObject]@{
          'Contribution Type'        = 'All Except Locked Awards'
          Shares                     = '=SUM({{Shares}}3:{{Shares}}5)'
          Value                      = '=SUM({{Value}}3:{{Value}}5)'
          'Cost Basis'               = '=SUM({{Cost Basis}}3:{{Cost Basis}}5)'
          'Own Costs'                = '=SUM({{Own Costs}}3:{{Own Costs}}5)'
          'Taxable Unrealized Gains' = '=SUM({{Taxable Unrealized Gains}}3:{{Taxable Unrealized Gains}}5)'
          'Real Unrealized Gains'    = '=SUM({{Real Unrealized Gains}}3:{{Real Unrealized Gains}}5)'
          'Estimated Tax'            = '=SUM({{Estimated Tax}}3:{{Estimated Tax}}5)'
          'Estimated Net Profit'     = '=SUM({{Estimated Net Profit}}3:{{Estimated Net Profit}}5)'
          'Percentage Gain'          = '=ROUND({{Estimated Net Profit}}{{index}} / {{Own Costs}}{{index}} * 100, 2)'
        }) | Out-Null

      return $items | ConvertTo-ItemsWithCellReferences -StartIndex 2
    }

    function New-TaxItems {
      return @(
        [PSCustomObject]@{
          Tax               = 'Income'
          'Rate Percentage' = 42
          Rate              = "=B2 / 100"
        },
        [PSCustomObject]@{
          Tax               = 'Capital Gains'
          'Rate Percentage' = 26.375
          Rate              = "=B3 / 100"
        }
      )
    }

    $outputFiles = [System.Collections.ArrayList]@()
    $generateOutputFilePath = $false
    if (-not $OutputFile) {
      $generateOutputFilePath = $true
    }
  }
  process {
    if (-not (Test-Path -Path $InputFile -PathType Leaf)) {
      Write-Error "Input file `"$([System.IO.Path]::GetFullPath($InputFile))`" does not exist."
      exit 1
    }

    if ($MyInvocation.ExpectingInput) {
      if ($PSBoundParameters.ContainsKey('OutputFile')) {
        Write-Error 'The OutputFile parameter cannot be specified when using pipeline input.'
        exit 2
      }
      $generateOutputFilePath = $true
    }

    if ($generateOutputFilePath) {
      Write-Verbose 'No OutputFile parameter provided. Constructing default path.'
      $directory = [System.IO.Path]::GetDirectoryName($InputFile)
      $filename = [System.IO.Path]::GetFileName($InputFile)
      $OutputFile = Join-Path -Path $directory -ChildPath "Enhanced-${filename}"
      Write-Verbose "Constructed output file path `"${OutputFile}`"."
    }

    Use-ImportExcelModuleRequiredVersion -ScriptBlock {
      $rowIndex = 6
      Write-Verbose "Reading Excel data starting with row ${rowIndex} from file `"${InputFile}`"."
      $inputData = Import-Excel -Path $InputFile -StartRow $rowIndex
      Write-Verbose "Imported $($inputData.Count) rows from `"${InputFile}`"."

      $internalItems = [System.Collections.ArrayList]@()
      foreach ($row in $inputData) {
        $rowIndex++
        $internalItems.Add((ConvertTo-InternalItem -ExcelRow $row -Index $rowIndex)) | Out-Null
      }
      $internalItems = $internalItems | Sort-Object -Property Date

      $excelFile = New-CleanFileStructure -FilePath $OutputFile
      $outputFiles.Add($excelFile) | Out-Null

      Write-Verbose 'Creating "Overview" worksheet.'
      $params = @{
        Path          = $excelFile
        WorksheetName = 'Overview'
        TableStyle    = 'Light1'
        FreezeTopRow  = $true
        AutoSize      = $true
      }
      New-OverviewItems -LastRowIndex ($internalItems.Count + 1) | Export-Excel @params

      Write-Verbose 'Creating "Tax Rates" worksheet.'
      $params = @{
        Path          = $excelFile
        WorksheetName = 'Tax Rates'
        TableStyle    = 'Light1'
        FreezeTopRow  = $true
        AutoSize      = $true
      }
      New-TaxItems | Export-Excel @params

      Write-Verbose 'Creating "Detailed Data" worksheet.'
      $params = @{
        Path          = $excelFile
        WorksheetName = 'Detailed Data'
        TableStyle    = 'Light1'
        FreezeTopRow  = $true
        AutoSize      = $true
      }
      $internalItems | ConvertTo-DetailedItems | Export-Excel @params

      Write-Verbose 'Formatting dates in the input data.'
      foreach ($item in $inputData) {
        $beginning = Get-Date '1899-12-30'
        foreach ($property in @('Allocation date', 'Available from', 'Expiry date')) {
          $item.$property = $beginning.AddDays($item.$property).ToString('yyyy-MM-dd')
        }
      }
      Write-Verbose 'Creating "Input Data" worksheet.'
      $params = @{
        Path          = $excelFile
        WorksheetName = 'Input Data'
        TableStyle    = 'Light1'
        FreezeTopRow  = $true
        AutoSize      = $true
        Show          = $true
      }
      $inputData | Export-Excel @params
    }
  }
  end {
    return $outputFiles
  }
}
