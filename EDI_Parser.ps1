# Define the paths for the input files
$numberFile = ""  # Path to the list of numbers
$dataFiles = @("")  # Data files

# Read the list of numbers from the first file
$numbers = (Get-Content $numberFile) -join ',' -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
$num = ""

    # Iterate over each data file
    foreach ($filePath in $dataFiles) {
        # Read the content of the current data file as a single string (raw)
        $fileContent = Get-Content -Path $filePath -Raw

        # Debug: Inspect the first few characters of the raw content
        Write-Host "Inspecting file: $filePath"
        Write-Host "First 100 characters of raw content:"
        Write-Host $fileContent.Substring(0, [Math]::Min(100, $fileContent.Length))

        # Use regex to find blocks between quotes (multiline content between quotes)
        # This pattern will match everything between two double quotes, including newlines
        $matches = [regex]::Matches($fileContent, '"(.*?)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)

        # Debug: Print how many matches were found
        Write-Host "Found blocks in the file."

        # Iterate through all matches (i.e., blocks between quotes)
        foreach ($match in $matches) {
            $block = $match.Groups[1].Value

            # Debug: Show the first 100 characters of the block to inspect it
            $num = ($block -split "`r`n" | Where-Object { $_ -match "REF\*ICN\*" }) | Select-Object -Unique
            $nums = $num.Substring(8,14)
            #Write-Host $nums
            
            # Check if the current block contains the REF*ICN* with the current number
            if ($numbers -contains $nums) {
                Write-Host "Inspecting block:"
                Write-Host $block.Substring(0, [Math]::Min(1000, $block.Length))

                # If found, set the found data to the matching block
                $outputFilePath = "<PATH>\$nums.good"
                Set-Content -Path $outputFilePath -Value $block
                Write-Host "Matched data for $nums has been written to $outputFilePath"

                # Remove the matched number from the numbers list
                $numbers = $numbers | Where-Object { $_ -ne $nums }

                # Update the input file by overwriting with remaining numbers
                Set-Content -Path $numberFile -Value ($numbers -join ',') -Encoding UTF8
                
            }
        }
    }

