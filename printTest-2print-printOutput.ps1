# Stop for all errors
$ErrorActionPreference = 'Stop'
#
# 2PRINT Standard Test
#
# This function is an explanation of the script, and can be accessed with -help	
function Display-Help() {
	Write-Output "
	This script will execute print commands for the test pdf files using 2PRINT
	Updated for 2PRINT v1.1.1

	This script requires PowerShell Core 6

	Usage: $(split-path $MyInvocation.PSCommandPath -Leaf) <2PRINT-installation-directory> <test-pdf-directory> <OPTIONAL-test-case-to-invoke>
			Remember to update the json file or the json test will fail

	Arguments:
		PRINT executable directory
		test pdf directory - Point this script at the test pdf file directory. It expects the provided test pdfs and json file.
		test case - by name

	It expects the printers in $printerList to be available to your computer:

	You can check which printers are currently added to your PC with this powershell command:
		Get-CimInstance -ClassName Win32_Printer | Format-Wide
		
	Available Test Cases:"
	Foreach ($test in $testCaseList) {
		Write-Output "		$test"
	}	
}

$testCaseList=@("Test-Duplex",
				"Test-Ranges",
				"Test-ShrinkToFit",
				"Test-Copies",
				"Test-Collate",
				"Test-Orientation",
				"Test-Default",
				"Test-Showtrays",
				"Test-TwoFiles",
				"Test-Control")

# $printerList = @("LBP653C/654C", "RICOH SP 4510DN", "Color Printer S3840cdn", "VersaLink C400", "Brother HL-L8360CDW series Printer")
$printerList = @("VersaLink C400")
# $printerList = @("RICOH SP 4510DN")
# $printerList = @("Brother HL-L8360CDW series Printer")
# $printerList = @("LBP653C/654C")
# $printerList = @("Color Printer S3840cdn")

#region Utility Functions
# Sends most print commands to the printer. Print-Default also send print commands
function Print-Output($testFile, $testPrinter, $param) {
	Write-Output "Printing $testFile to $testPrinter with parameters: $param"
	&$printExe "--printer=$testPrinter" $param $testFile
	Write-Output ""
}

# Print output with two parameters, or prints two pdf files with one command
function Print-OutputTwo($testFile, $testPrinter, $param, $var) {
	Write-Output "Printing to $testPrinter with multiple files or parameters"
	&$printExe "--printer=$testPrinter" $param $var $testFile
	Write-Output ""
}

# Print using a single parameter - control (json) file
function Print-ControlFile($param) {
	Write-Output "Printing to using parameter: $param"
	&$printExe $param
	Write-Output ""
}

# Checks if the recieved test printer exists. If true, then sets testprinter as default printer and passes testfile to Print-Output
function Print-Default($testFile, $testPrinter) {
	if (!(Get-CimInstance -ClassName Win32_Printer -Filter "Name = '$($testPrinter)'")){
		Write-Output "Printer not found: $testPrinter"
		Write-Output "    Skipping $testPrinter default printer test..."
	}
	else {
		Write-Output "Setting $testprinter as the default printer..."
		(New-Object -ComObject WScript.Network).SetDefaultPrinter("$testPrinter")
		Write-Output "Printing $testFile the default printer - should print from: $testPrinter"
		&$printExe $testFile
	}
	Write-Output ""
}

# Calls 2print against a printer with a single parameter.
function Print-ParamOnly($testPrinter, $param) {
	Write-Output "Calling 2PRINT.exe $param for Printer: $testPrinter"
	&$printExe "--printer=$testPrinter" $param
	Write-Output ""
}

# Checks if the recieved test pdf exists
function Check-Testfile($testFile) {
	if (!(Test-Path -Path $testFile -PathType leaf)) {
		Write-Output "Failed to find the test file: $testFile"
		Exit
	}
}

function Display-Usage() {
	Write-Output "Usage: $(split-path $MyInvocation.PSCommandPath -Leaf)  <2PRINT-installation-directory> <test-pdf-directory> <OPTIONAL-test-case-to-invoke>
Use -help for info and test cases, e.g.: $(split-path $MyInvocation.PSCommandPath -Leaf) -help
	Remember to update the json file or the json test will fail"
}
#endregion

#region Test Cases
#Duplex printing
function Test-Duplex() {
	$duplexTest = @("duplex-default.pdf", "duplex-long.pdf", "duplex-short.pdf", "duplex-off.pdf")
	Foreach ($testCase in $duplexTest) {
		$testFile="$inDir\duplex\$testCase"
		Check-Testfile $testFile
		Foreach ($printer in $printerList) {
			if ($printer -eq "VersaLink C400") {
				Write-Output "Skipping $testCase for VersaLink C400 - duplex unsupported"
				break
			}
			else {
				switch ($testCase) {
					"duplex-default.pdf" {Print-Output $testFile $printer "--duplex=default"; break}
					"duplex-long.pdf" {
						if ($printer -eq "Brother HL-L8360CDW series Printer") {
							Write-Output "Skipping $testCase for Brother HL-L8360CDW series Printer - duplex long unsupported"
							break
						}
						else {
							Print-Output $testFile $printer "--duplex=long"; break
						}
					}
					"duplex-short.pdf" {Print-Output $testFile $printer "--duplex=short"; break}
					"duplex-off.pdf" {Print-Output $testFile $printer "--duplex=off"; break}
				}
			}
		}
	}
	Write-Output ""
	Write-Output ""
}

function Test-Ranges() {
	$rangesTest = @("ranges-6to8.pdf", "ranges-all.pdf", "ranges-even.pdf", "ranges-odd.pdf")
	Foreach ($testCase in $rangesTest) {
		$testFile="$inDir\ranges\$testCase"
		Check-Testfile $testFile
		Foreach ($printer in $printerList) {
			switch ($testCase) {
				"ranges-6to8.pdf" {Print-Output $testFile $printer "--range=2-4"; break}
				"ranges-all.pdf" {Print-Output $testFile $printer "--range=all"; break}
				"ranges-even.pdf" {Print-Output $testFile $printer "--range=even"; break}
				"ranges-odd.pdf" {Print-Output $testFile $printer "--range=odd"; break}
			}
		}
	}
	Write-Output ""
	Write-Output ""
}

function Test-ShrinkToFit() {
	$shrinkToFitTest = @("shrinkTrue.pdf", "shrinkFalse.pdf")
	Foreach ($testCase in $shrinkToFitTest) {
		$testFile="$inDir\shrink\$testCase"
		Check-Testfile $testFile
		Foreach ($printer in $printerList) {
			switch ($testCase) {
				"shrinkTrue.pdf" {Print-Output $testFile $printer "--shrinktofit=true"; break}
				"shrinkFalse.pdf" {Print-Output $testFile $printer "--shrinktofit=false"; break}
			}
		}
	}
	Write-Output ""
	Write-Output ""
}

function Test-Copies() {
	$ncopiesTest = "3copies.pdf"
	Foreach ($testCase in $ncopiesTest) {
		$testFile="$inDir\$testCase"
		Check-Testfile $testFile
		Foreach ($printer in $printerList) {
			Print-Output $testFile $printer "--copies=3"
		}
	}
	Write-Output ""
	Write-Output ""
}

function Test-Collate() {
	$collateTest = @("collate.pdf", "nocollate.pdf")
	Foreach ($testCase in $collateTest) {
		$testFile="$inDir\collate\$testCase"
		Check-Testfile $testFile
		Foreach ($printer in $printerList) {
			if ($printer -eq "RICOH SP 4510DN") {
				Write-Output "Skipping $testCase for RICOH SP 4510DN - collate=false unsupported, default is true"
				break
			}
			else {
				switch ($testCase) {
					"collate.pdf" {Print-OutputTwo $testFile $printer "--copies=2" "--collate=true"; break}
					"nocollate.pdf" {Print-OutputTwo $testFile $printer "--copies=2" "--collate=false"; break}
				}
			}
		}
	}
	Write-Output ""
	Write-Output ""
}

function Test-Orientation() {
	$orientationTest = @("orientation-default.pdf", "orientation-landscape.pdf", "orientation-portrait.pdf")
	Foreach ($testCase in $orientationTest) {
		$testFile="$inDir\orientation\$testCase"
		Check-Testfile $testFile
		Foreach ($printer in $printerList) {
			switch ($testCase) {
				"orientation-default.pdf" {Print-Output $testFile $printer "--orientation=default"; break}
				"orientation-landscape.pdf" {Print-Output $testFile $printer "--orientation=landscape"; break}
				"orientation-portrait.pdf" {Print-Output $testFile $printer "--orientation=portrait"; break}
			}
		}
	}
	Write-Output ""
	Write-Output ""
}

# function Test-Tofile() {
	# $orientationTest = @("tofile.pdf")
	# Foreach ($testCase in $orientationTest) {
		# $testFile="$inDir\$testCase"
		# $param="-tofile=$($args[0])\tofile.ps"
		# Check-Testfile $testFile
		# Write-Output "Printing $testFile to ps with parameters: $param"
		# &$printExe $testFile $param
	# }
# }

function Test-Showtrays() {
	Foreach ($printer in $printerList) {
		$param="--showtrays"
		Write-Output "Testing -showtrays with printer: $printer"
		Print-ParamOnly $printer $param
	}
	Write-Output ""
	Write-Output ""
}

function Test-Default() {
	# Save current default printer
	$currentDefaultPrinter=$(Get-CimInstance -ClassName Win32_Printer -Filter "DEFAULT=$true" | Select-Object -ExpandProperty Name)
	# Set each printer as default, then pass the testfile and test printer Print-Default
	Foreach ($printer in $printerList) {
		switch ($printer) {
			"Brother HL-L8360CDW series Printer" {
				$testFile="$inDir\default\default-brother.pdf"; Check-Testfile $testFile; 
				Print-Default $testFile "Brother HL-L8360CDW series Printer"; break
			}
			"LBP653C/654C" {
				$testFile="$inDir\default\default-lbp653c.pdf"; Check-Testfile $testFile; 
				Print-Default $testFile "LBP653C/654C"; break
			}
			"RICOH SP 4510DN" {
				$testFile="$inDir\default\default-ricoh.pdf"; Check-Testfile $testFile; 
				Print-Default $testFile "RICOH SP 4510DN"; break
			}
			"Color Printer S3840cdn" {
				$testFile="$inDir\default\default-s3840cdn.pdf"; Check-Testfile $testFile; 
				Print-Default $testFile "Color Printer S3840cdn"; break
			}
			"VersaLink C400" {
				$testFile="$inDir\default\default-versalink.pdf"; Check-Testfile $testFile; 
				Print-Default $testFile "VersaLink C400"; break
			}
		}
		# After the test, restore the default printer prior to continuing
		$tempDefaultPrinter=$(Get-CimInstance -ClassName Win32_Printer -Filter "DEFAULT=$true" | Select-Object -ExpandProperty Name)
		if ($tempDefaultPrinter -ne $currentDefaultPrinter) {
			Write-Output "Restoring default printer"
			(New-Object -ComObject WScript.Network).SetDefaultPrinter("$currentDefaultPrinter")
		}
	}
	Write-Output ""
	Write-Output ""
}

function Test-TwoFiles() {
	$twofileTestA="$inDir\multiplePdf\multiplePdf-fileA.pdf"
	$twofileTestB="$inDir\multiplePdf\multiplePdf-fileB.pdf"
	Check-Testfile $twofileTestA
	Check-Testfile $twofileTestB
	Foreach ($printer in $printerList) {
		Print-OutputTwo $twofileTestA $printer "--copies=2" $twofileTestB
	}
	Write-Output ""
	Write-Output ""
}

function Test-Control() {
	$testFile="$inDir\json\jsonTest.json"
	Check-Testfile $testFile
	Print-ControlFile "--control=$testFile"
	Write-Output ""
	Write-Output ""
}

#endregion

# Manage args
if (($args.length -lt 1) -or ($args.length -gt 3)) {
	Display-Usage
	Exit
}

if ($args.length -eq 1) {
	if ($args[0] -eq "-help") {
		Display-Help
	}
	else {
		Display-Usage
	}
	Exit
}

if ($args.length -eq 3) {
	if (!($testCaseList.Contains("$($args[2])"))) {
		Write-Output "$($args[2]) is not a valid test case"
		Display-Usage
		Exit
	}
	else {
		$oneTestCase="$($args[2])"
	}
}

$printExe="$($args[0])\2print.exe"
$inDir="$($args[1])"

Write-Output "Starting 2PRINT Test"
Write-Output "-----------------------"

# Check available printers
Write-Output "Checking printers"
Foreach ($i in $printerList) {
	Write-Output "Checking $i..."
	if (!(Get-CimInstance -ClassName Win32_Printer -Filter "Name = '$($i)'")) {
		Write-Output "Printer not found: $i"
		Exit
	}
}
Write-Output "All printers found"
Write-Output ""
Write-Output ""

# Check 2PRINT installation directory
if (!(Test-Path -Path $printExe -PathType leaf)) {
	Write-Output "Cannot find 2PRINT in $($args[0])"
	Exit
}
 
# Check input directory exists
if (!(Test-Path -Path $inDir -PathType Container)) {
	Write-Output "Cannot find input file directory at $($inDir)"
	Exit
}
else {
	# Check that input directory has pdf files somewhere
	if ((Get-ChildItem -Path $inDir -Recurse -Filter *.pdf).count -eq 0){
		Write-Output "No PDF files located at input directory"
		Exit
	}
	else {
		Write-Output "Executing 2PRINT Test..."
		Write-Output ""
		# If one test is specified then only execute that test
		if ($args.length -eq 3) {
			&$oneTestCase
		}
		# Otherwise execute all tests
		else {
			Foreach ($test in $testCaseList) {
				&$test
			}
		}
	}
}
