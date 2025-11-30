<#
   Copyright 2025 Adam Gabryś

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

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
- Own Costs: The employee's costs for shares received through the employer's Company Match.
- Percentage Gain: Estimated Net Profit expressed as a percentage of Own Costs.
- Plan: The name of the portfolio plan.
- Real Unrealized Gains: Actual unrealized gains after deducting Own Costs.
- Shares: The number of allocated shares.
- Taxable Unrealized Gains: Gains subject to taxation, calculated as Value − Cost Basis.
- Value: The market value of the allocated shares, calculated as Shares × Market Price.

The function ensures the required version of the ImportExcel module (7.8.10) is installed and imported
during execution, restoring any previously loaded versions afterward.

It returns a list of PSCustomObject objects, where each object stores two properties:
1. InputFile – absolute path to the source Excel report,
2. OutputFile – absolute path to the created enhanced Excel report.

.PARAMETER InputFiles
Paths to the source Excel files containing portfolio details.
All generated report files will use the same names as the source files, prefixed with "Enhanced-".

.PARAMETER InputFile
Path to a single source Excel file containing portfolio details.

.PARAMETER OutputDir
Path to the directory where the enhanced Excel report(s) will be created.
If not specified, the reports will be saved in the current directory.

.PARAMETER OutputFile
Path where the enhanced Excel report will be created.
If not specified, the output file will default to "Enhanced-" followed by the input file name.

.PARAMETER IncomeTax
Specifies the income tax percentage used to calculate the employee's Own Costs for shares
purchased by the employer as part of the Company Match.
This value must be a double between 0.0 and 100.0 (inclusive).
The default is 42.0.

.PARAMETER CapitalGainsTax
Specifies the capital gains tax percentage to apply.
This value must be a double between 0.0 and 100.0 (inclusive).
The default is 26.375.

.PARAMETER Open
Specifies whether the enhanced Excel report(s) should be opened automatically after creation.

.INPUTS
System.String[]. An array of paths to the source Excel files.

.INPUTS
None. When the -InputFile parameter is used, no values are accepted from the pipeline.

.OUTPUTS
System.Collections.IList<PSCustomObject>. A list of PSCustomObject objects, where each object stores two properties:
1. InputFile – absolute path to the source Excel report,
2. OutputFile – absolute path to the created enhanced Excel report.

.EXAMPLE
'Portfolio1.xlsx' | New-EnhancedEquatePlusPortfolioDetails

Accepts an input file from the pipeline, applies an income tax of 42% and a capital gains tax of 26.375%,
and saves the report in the current directory as Enhanced-Portfolio1.xlsx.

.EXAMPLE
'Portfolio1.xlsx', 'Portfolio2.xlsx' | New-EnhancedEquatePlusPortfolioDetails -OutputDir 'C:\reports' -IncomeTax 35 -CapitalGainsTax 20

Accepts multiple input files from the pipeline, applies an income tax of 35% and capital gains tax of 20%,
and saves enhanced reports in the C:\reports directory.

.EXAMPLE
New-EnhancedEquatePlusPortfolioDetails -InputFile 'Portfolio.xlsx' -OutputFile 'DetailedPortfolio.xlsx' -IncomeTax 45 -CapitalGainsTax 30

Generates a detailed Excel report from Portfolio.xlsx, applies an income tax of 45% and a capital gains tax of 30%,
and saves the report as DetailedPortfolio.xlsx.

.EXAMPLE
New-EnhancedEquatePlusPortfolioDetails -InputFile 'Portfolio.xlsx' -Verbose

Shows detailed progress for module loading, row processing, and worksheet creation.

.NOTES
Author:  Adam Gabryś
Date:    2025-11-30
Version: 0.5.0
License: Apache-2.0

.LINK
https://github.com/agabrys/equateplus-portfolio-details-enhancer
#>
function New-EnhancedEquatePlusPortfolioDetails {
  [CmdletBinding(SupportsShouldProcess = $true)]
  param (
    [Parameter(ParameterSetName = 'Batch', Mandatory = $true, ValueFromPipeline = $true)]
    [string[]]$InputFiles,
    [Parameter(ParameterSetName = 'Batch')]
    [string]$OutputDir,

    [Parameter(ParameterSetName = 'Single', Mandatory = $true)]
    [string]$InputFile,
    [Parameter(ParameterSetName = 'Single')]
    [string]$OutputFile,

    [Parameter()]
    [ValidateRange(0.0, 100.0)]
    [double]$IncomeTax = 42.0,
    [Parameter()]
    [ValidateRange(0.0, 100.0)]
    [double]$CapitalGainsTax = 26.375,

    [Parameter()]
    [switch]$Open
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
        Write-Verbose "Required version ${version} of the ${name} module is not installed."
        if ($PSCmdlet.ShouldProcess("Version '${version}' of module '${name}'", 'Install-Module')) {
          Install-Module -Name $name -RequiredVersion $version -Scope CurrentUser -Force -AllowClobber
        } else {
          if ($WhatIfPreference) {
            Write-Verbose "Skipping execution, as no processing can be done without the version ${version} of the ${name} module."
          } else {
            Write-Error "No processing can be done without the version ${version} of the ${name} module."
          }
          return
        }
      } else {
        Write-Verbose "Required version ${version} of the ${name} module is installed."
      }

      $originalModule = Get-Module -Name $name
      if ($null -eq $originalModule) {
        Write-Verbose "${name} module is not currently imported."
      } else {
        Write-Verbose "The currently imported version of the ${name} module is $($originalModule.Version)."
      }

      Import-Module -Name $name -RequiredVersion 7.8.9

      #TODO

      $result = $null
      try {
        if ($null -ne $originalModule -and $originalModule.Version -ne $version) {
          Write-Verbose "Removing the currently imported version $($originalModule.Version) of the ${name} module as version ${version} is needed."
          Remove-Module -Name $name
        }
        if ($null -eq $originalModule -or $originalModule.Version -ne $version) {
          Write-Verbose "Importing version ${version} of the ${name} module."
          Import-Module -Name $name -RequiredVersion $version
        }

          #$result = & $ScriptBlock

      } finally {
        if ($null -eq $originalModule -or $originalModule.Version -ne $version) {
          Write-Verbose "Restoring the ${name} module to its state before the scriptlet was executed."
          Write-Verbose "Removing the currently imported version ${version} of the ${name} module."
          Remove-Module -Name $name
          if ($null -ne $originalModule) {
            Write-Verbose "Importing version $($originalModule.Version) of the ${name} module."
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

    function Initialize-CleanFilePath {
      param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
      )
      $parentDir = [System.IO.Path]::GetDirectoryName($FilePath)
      if (-not (Test-Path -Path $parentDir)) {
        Write-Verbose "Creating directory `"${parentDir}`"."
        New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
      }
      if (Test-Path -Path $FilePath) {
        Write-Verbose "Removing file `"${FilePath}`"."
        Remove-Item -Path $FilePath -Force | Out-Null
      }
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
      param(
        [Parameter(Mandatory = $true)]
        [double]$IncomeTax,
        [Parameter(Mandatory = $true)]
        [double]$CapitalGainsTax
      )
      Write-Verbose "Tax rates: income = ${IncomeTax}%, capital gains = ${CapitalGainsTax}%."
      return @(
        [PSCustomObject]@{
          Tax               = 'Income'
          'Rate Percentage' = $IncomeTax
          Rate              = "=B2 / 100"
        },
        [PSCustomObject]@{
          Tax               = 'Capital Gains'
          'Rate Percentage' = $CapitalGainsTax
          Rate              = "=B3 / 100"
        }
      )
    }

    $currentInputFiles = [System.Collections.ArrayList]@()

    if ($OutputDir) {
      $OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
    }
    if ($OutputFile) {
      $OutputFile = [System.IO.Path]::GetFullPath($OutputFile)
    }
  }
  process {
    $files = if ($PSCmdlet.ParameterSetName -eq 'Batch') { $InputFiles } else { @($InputFile) }
    foreach ($file in $files) {
      $file = [System.IO.Path]::GetFullPath($file)
      if (-not (Test-Path -Path $file -PathType Leaf)) {
        Write-Error "Input file `"${file}`" does not exist."
        exit 1
      }
      $currentInputFiles.Add([System.IO.Path]::GetFullPath($file)) | Out-Null
    }
  }
  end {
    $output = [System.Collections.ArrayList]@()

    Use-ImportExcelModuleRequiredVersion -ScriptBlock {
      foreach ($currentInputFile in $currentInputFiles) {
        Write-Verbose "Processing file `"${currentInputFile}`"."

        $currentOutputFile = $OutputFile
        if (-not $OutputFile) {
          $directory = if ($OutputDir) { $OutputDir } else { [System.IO.Path]::GetDirectoryName($currentInputFile) }
          $filename = [System.IO.Path]::GetFileName($currentInputFile)
          $currentOutputFile = Join-Path -Path $directory -ChildPath "Enhanced-${filename}"
          Write-Verbose "Constructed output file path `"${currentOutputFile}`"."
        }

        $rowIndex = 6
        Write-Verbose "Reading Excel data starting with row ${rowIndex} from file `"${currentInputFile}`"."
        $inputData = Import-Excel -Path $currentInputFile -StartRow $rowIndex
        Write-Verbose "Imported $($inputData.Count) rows from `"${currentInputFile}`"."

        $internalItems = [System.Collections.ArrayList]@()
        foreach ($row in $inputData) {
          $rowIndex++
          $internalItems.Add((ConvertTo-InternalItem -ExcelRow $row -Index $rowIndex)) | Out-Null
        }
        $internalItems = $internalItems | Sort-Object -Property Date

        Initialize-CleanFilePath -FilePath $currentOutputFile
        $output.Add([PSCustomObject]@{
            InputFile  = $currentInputFile
            OutputFile = $currentOutputFile
          }) | Out-Null

        if ($PSCmdlet.ShouldProcess($currentOutputFile, 'Create File')) {
          Write-Verbose 'Creating "Overview" worksheet.'
          $params = @{
            Path          = $currentOutputFile
            WorksheetName = 'Overview'
            TableStyle    = 'Light1'
            FreezeTopRow  = $true
            AutoSize      = $true
          }
          New-OverviewItems -LastRowIndex ($internalItems.Count + 1) | Export-Excel @params

          Write-Verbose 'Creating "Tax Rates" worksheet.'
          $params = @{
            Path          = $currentOutputFile
            WorksheetName = 'Tax Rates'
            TableStyle    = 'Light1'
            FreezeTopRow  = $true
            AutoSize      = $true
          }
          New-TaxItems -IncomeTax $IncomeTax -CapitalGainsTax $CapitalGainsTax | Export-Excel @params

          Write-Verbose 'Creating "Detailed Data" worksheet.'
          $params = @{
            Path          = $currentOutputFile
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
            Path          = $currentOutputFile
            WorksheetName = 'Input Data'
            TableStyle    = 'Light1'
            FreezeTopRow  = $true
            AutoSize      = $true
            Show          = $Open
          }
          $inputData | Export-Excel @params

          Write-Verbose "Saved file `"${currentOutputFile}`"."
        }
      }
    }

    return $output
  }
}
