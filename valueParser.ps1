# Define input and output files
$numberFile = "<filepath>"
$dataFiles = @(
   "<filepath>"
)
$outputFile = "<filepath>"

# Read numbers from the first file and split by comma
$numbers = (Get-Content $numberFile) -join ',' -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

# Initialize the output content
$outputContent = @()

# Process each data file
foreach ($dataFile in $dataFiles) {
    $lines = Get-Content $dataFile
    
    foreach ($line in $lines) {
        # Use a regex to extract the relevant number (after the third whitespace)
        $matches = [regex]::Matches($line, '\S+')
        
        if ($matches.Count -gt 3) {
            # Get the number (4th match) and trim it
            $number = $matches[3].Value.Trim()

            # Only compare the first 14 characters
            $numberToCompare = $number.Substring(0, [Math]::Min(14, $number.Length))

            # Check if this number (first 14 chars) is in the list of numbers
            if ($numbers -contains $numberToCompare) {
                $outputContent += $line
                Write-Host "Match found: $line"  # Debug output

                # Remove the matched number from the numbers list to avoid duplicates
                $numbers = $numbers | Where-Object { $_ -ne $numberToCompare }
            }
        }
    }
}

# Write the output content to the output file
if ($outputContent.Count -gt 0) {
    $outputContent | Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host "Processing complete. Results saved to $outputFile"
} else {
    Write-Host "No matching lines found."
}
