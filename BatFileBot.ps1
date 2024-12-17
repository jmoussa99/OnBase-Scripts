# Ensure that the required Add-Type is present for mouse click simulation
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class MouseOperations
{
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
    
    public const int MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const int MOUSEEVENTF_LEFTUP = 0x0004;
    public const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
    public const int MOUSEEVENTF_RIGHTUP = 0x0010;
    
    // Simulate mouse click
    public static void Click(int x, int y)
    {
        SetCursorPos(x, y);
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
    }
}
"@ 

# Manually load System.Windows.Forms assembly to use SendKeys
Add-Type -AssemblyName "System.Windows.Forms"

# Function to send keystrokes
function SendKeys {
    param([string]$keys)
    [System.Windows.Forms.SendKeys]::SendWait($keys)
}

$servers = @(
    @{ Name = "Server61"; X = 92; Y = 88 },
    @{ Name = "Server63"; X = 92; Y = 103 },  # Replace with actual coordinates
    @{ Name = "Server64"; X = 92; Y = 116 },
    @{ Name = "Server65"; X = 92; Y = 136 },
    @{ Name = "Server66"; X = 92; Y = 148 },
    @{ Name = "Server67"; X = 92; Y = 163 },
    @{ Name = "Server68"; X = 92; Y = 182 },
    @{ Name = "Server69"; X = 92; Y = 196 },
    @{ Name = "Server70"; X = 92; Y = 213 },
    @{ Name = "Server71"; X = 92; Y = 228 },
    @{ Name = "Server72"; X = 92; Y = 244 },
    @{ Name = "Server73"; X = 92; Y = 262 },
    @{ Name = "Server74"; X = 92; Y = 277 },
    @{ Name = "Server75"; X = 92; Y = 292 },
    @{ Name = "Server76"; X = 92; Y = 307 },
    @{ Name = "Server77"; X = 92; Y = 327 },
    @{ Name = "Server78"; X = 92; Y = 340 },
    @{ Name = "Server79"; X = 92; Y = 360 },
    @{ Name = "Server80"; X = 92; Y = 377 },
    @{ Name = "Server81"; X = 92; Y = 391 },
    @{ Name = "Server82"; X = 92; Y = 406 },
    @{ Name = "Server83"; X = 92; Y = 421 },
    @{ Name = "Server84"; X = 92; Y = 440 },
    @{ Name = "Server85"; X = 92; Y = 451 },
    @{ Name = "Server86"; X = 92; Y = 467 },
    @{ Name = "Server87"; X = 92; Y = 483 },
    @{ Name = "Server88"; X = 92; Y = 500 },
    @{ Name = "Server89"; X = 92; Y = 518 },
    @{ Name = "Server90"; X = 92; Y = 532 },
    @{ Name = "Server91"; X = 92; Y = 548 },
    @{ Name = "Server92"; X = 92; Y = 564 },
    @{ Name = "Server93"; X = 92; Y = 580 },
    @{ Name = "Server94"; X = 92; Y = 598 },
    @{ Name = "Server95"; X = 92; Y = 611 },
    @{ Name = "Server96"; X = 92; Y = 628 },
    @{ Name = "Server97"; X = 92; Y = 643 },
    @{ Name = "Server98"; X = 92; Y = 660 },
    @{ Name = "Server99"; X = 92; Y = 677 },
    @{ Name = "Server100"; X = 92; Y = 692 },
    @{ Name = "Server101"; X = 92; Y = 710 },
    @{ Name = "Server102"; X = 92; Y = 725 },
    @{ Name = "Server103"; X = 92; Y = 741 },
    #@{ Name = "Server104"; X = 92; Y = 756 },
    @{ Name = "Server105"; X = 92; Y = 773 },
    @{ Name = "Server106"; X = 92; Y = 790 },
    @{ Name = "Server107"; X = 92; Y = 804 },
    @{ Name = "Server108"; X = 92; Y = 819 },
    @{ Name = "Server109"; X = 92; Y = 835 },
    @{ Name = "Server110"; X = 92; Y = 851 },
    @{ Name = "Server111"; X = 92; Y = 102 },
    @{ Name = "Server112"; X = 92; Y = 116 },
    @{ Name = "Server113"; X = 92; Y = 136 },
    @{ Name = "Server114"; X = 92; Y = 147 },
    @{ Name = "Server115"; X = 92; Y = 163 },
    @{ Name = "Server116"; X = 92; Y = 180 },
    @{ Name = "Server117"; X = 92; Y = 197 },
    @{ Name = "Server118"; X = 92; Y = 212 },
    @{ Name = "Server119"; X = 92; Y = 228 },
    @{ Name = "Server120"; X = 92; Y = 245 },
    @{ Name = "Server121"; X = 92; Y = 261 },
    @{ Name = "Server122"; X = 92; Y = 275 },
    @{ Name = "Server123"; X = 92; Y = 291 },
    @{ Name = "Server124"; X = 92; Y = 309 },
    @{ Name = "Server125"; X = 92; Y = 326 },
    @{ Name = "Server126"; X = 92; Y = 340 }, 
    @{ Name = "Server127"; X = 92; Y = 355 }, 
    @{ Name = "Server128"; X = 92; Y = 374 },
    @{ Name = "Server129"; X = 92; Y = 390 },
    @{ Name = "Server130"; X = 92; Y = 406 },
    @{ Name = "Server131"; X = 92; Y = 420 },
    @{ Name = "Server132"; X = 92; Y = 437 },
    @{ Name = "Server133"; X = 92; Y = 453 },
    @{ Name = "Server134"; X = 92; Y = 471 },
    @{ Name = "Server135"; X = 92; Y = 484 },
    @{ Name = "Server136"; X = 92; Y = 500 },
    @{ Name = "Server137"; X = 92; Y = 517 },
    @{ Name = "Server138"; X = 92; Y = 533 },
    @{ Name = "Server139"; X = 92; Y = 550 },
    @{ Name = "Server140"; X = 92; Y = 565 },
    @{ Name = "Server141"; X = 92; Y = 580 },
    @{ Name = "Server142"; X = 92; Y = 596 },
    @{ Name = "Server143"; X = 92; Y = 614 },
    @{ Name = "Server144"; X = 92; Y = 627 },
    @{ Name = "Server145"; X = 92; Y = 642 },#>
    @{ Name = "Server146"; X = 92; Y = 659 },
    @{ Name = "Server147"; X = 92; Y = 675 },
    @{ Name = "Server148"; X = 92; Y = 691 },
    @{ Name = "Server149"; X = 92; Y = 707 },
    @{ Name = "Server150"; X = 92; Y = 723 },
    @{ Name = "Server151"; X = 92; Y = 739 },
    @{ Name = "Server152"; X = 92; Y = 755 },
    @{ Name = "Server153"; X = 92; Y = 771 },
    @{ Name = "Server154"; X = 92; Y = 787 },
    @{ Name = "Server155"; X = 92; Y = 803 },
    @{ Name = "Server156"; X = 92; Y = 820 },
    @{ Name = "Server157"; X = 92; Y = 835 },
    @{ Name = "Server158"; X = 92; Y = 851 },
    @{ Name = "Server159"; X = 92; Y = 868 },
    @{ Name = "Server159"; X = 92; Y = 884 }
)

# Coordinates for digits (example values, you need to determine them based on your setup)
$digitCoordinates = @{
    "0" = [Tuple]::Create(1066, 720)
    "9" = [Tuple]::Create(988, 720)
    "8" = [Tuple]::Create(905, 720)
    "7" = [Tuple]::Create(843, 720)
    "6" = [Tuple]::Create(766, 720)
    "5" = [Tuple]::Create(683, 720)
    "4" = [Tuple]::Create(612, 720)
    "3" = [Tuple]::Create(530, 720)
    "2" = [Tuple]::Create(447, 720)
    "1" = [Tuple]::Create(383, 720)
}

# Function to click on a number based on its digit
function Click-Number {
    param (
        [string]$digit
    )
    
    if ($digitCoordinates.ContainsKey($digit)) {
        $coordinates = $digitCoordinates[$digit]
        [MouseOperations]::Click($coordinates.Item1, $coordinates.Item2)
    } else {
        Write-Host "Invalid digit: $digit"
    }
}

# Specify the folder path removed for privacy
$eastPath = ""
$westPath = ""

# Create an empty list to store the numbers

function GetNumbers{ param ([string] $folderPath)
    $numberList = @()
    # Get all files in the folder
    Get-ChildItem -Path $folderPath -File | ForEach-Object {
        # Get the file name without the extension
        $fileName = $_.BaseName
    
        # Split the filename by dashes
        $parts = $fileName -split ' - '

        # Check if there are enough parts (at least 5 parts expected for the example format)
        if ($parts.Length -ge 5) {
            # Get the number beween the first and second dashes (this is the 4th part from the end)
            $number = $parts[-2] # This works for cases like "1414" in the example filename

            # Add the number to the list
            $numberList += $number
        }
    }

    # Remove duplicates from the list
    return $numberList | Sort-Object -Unique
}

$uniqueNumbers = GetNumbers -folderPath $eastPath
$i = 10
# Loop through each server and automate the login process
foreach ($server in $servers) {
    
    if($server.Name -eq "server111"){
        [MouseOperations]::Click(30, 71)
    }
    Write-Host $server.Name

     if($server.Name -eq "server147"){
        $uniqueNumbers = GetNumbers -folderPath $westPath
        $i = 0
    }
    # Simulate double-clicking on the server (coordinates of the server in RDCM)
    [MouseOperations]::Click($server.X, $server.Y)
    [MouseOperations]::Click($server.X, $server.Y)
    Start-Sleep -Milliseconds 3500  # Pause to allow RDCM to open the connection
    [MouseOperations]::Click(2163, 300)
    # Navigate to the coordinates 912, 589 and click before entering the password
    [MouseOperations]::Click(912, 529)
    Start-Sleep -Milliseconds 1000 

    [System.Windows.Forms.SendKeys]::SendWait("^{V}")
    # Send password and hit Enter
    #SendKeys $password
    Start-Sleep -Milliseconds 1000  # Pause to allow password to be entered
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds 4500  # Pause to ensure server login starts
    
    
    # Press Enter twice as described
    [MouseOperations]::Click(877, 776)
    Start-Sleep -Milliseconds 9000
    [MouseOperations]::Click(1199, 625)
    Start-Sleep -Seconds 9  # Wait before continuing to the next server

    # Open File Explorer by clicking on the taskbar icon (coordinates 220, 990)
    [MouseOperations]::Click(253, 1019)
    Start-Sleep -Seconds 3  # Wait for File Explorer to open

    #Click on EAST Batch
    [MouseOperations]::Click(607, 460) 
    Start-Sleep -Seconds 4
    
    [MouseOperations]::Click(340, 1018)
    Start-Sleep -Seconds 1
      
    #click on search
    [MouseOperations]::Click(1436, 307)
    Start-Sleep -Seconds 2
   

    #type number
    $numberStr = $uniqueNumbers[$i].ToString()
    foreach ($digit in $numberStr.ToCharArray()) {
        Click-Number -digit $digit
        Start-Sleep -Milliseconds 300  # Pause between clicks for visibility
    }
    Start-Sleep -Seconds 2
    $i = $i + 1

    #run bat files
    [MouseOperations]::Click(1272, 841)
    Start-Sleep -Seconds 2

    #click on file
    [MouseOperations]::Click(815, 363)
    Start-Sleep -Seconds 1

    #highlight all
    [MouseOperations]::Click(308, 959)
    Start-Sleep -Seconds 1
    [MouseOperations]::Click(382, 844)
    Start-Sleep -Seconds 1

    #run bat files
    [MouseOperations]::Click(1272, 841)
    Start-Sleep -Seconds 1

    
    Start-Sleep -Seconds 2  # Wait for execution to start (adjust as needed)
}

Write-Host "Bot automation completed!"

