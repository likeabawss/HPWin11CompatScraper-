# HP Windows 11 Compatibility Scraper - Robust Selenium Version
# This script handles Chrome detection and ChromeDriver setup automatically

param(
    [switch]$ShowProgress,
    [switch]$HeadlessBrowser,
    [int]$WaitTimeSeconds = 15
)

# URL to scrape
$url = "https://support.hp.com/us-en/document/ish_4890350-4890415-16"


Write-Host "HP Windows 11 Compatibility Scraper - Robust Selenium" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
Write-Host ""

# Always delete previous output files on start
$htmlFile = "selenium_page_source.html"
$jsonFile = "HP_Win11_Compatibility.json"
if (Test-Path $htmlFile) { Remove-Item $htmlFile -Force -ErrorAction SilentlyContinue }
if (Test-Path $jsonFile) { Remove-Item $jsonFile -Force -ErrorAction SilentlyContinue }


# Function to download and install Chrome
function Install-GoogleChrome {
    Write-Host "Google Chrome not found. Downloading and installing..." -ForegroundColor Yellow
    
    try {
        # Create temp directory for Chrome installer
        $tempDir = "$env:TEMP\ChromeInstaller"
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }
        
        # Download Chrome installer
        $chromeInstallerUrl = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
        $installerPath = "$tempDir\chrome_installer.exe"
        
        Write-Host "Downloading Chrome installer..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $chromeInstallerUrl -OutFile $installerPath
        
        if (Test-Path $installerPath) {
            Write-Host "Installing Chrome..." -ForegroundColor Cyan
            
            # Run installer silently
            $installProcess = Start-Process -FilePath $installerPath -ArgumentList "/silent", "/install" -Wait -PassThru
            
            if ($installProcess.ExitCode -eq 0) {
                Write-Host "Chrome installed successfully!" -ForegroundColor Green
                
                # Wait a moment for installation to complete
                Start-Sleep -Seconds 3
                
                # Try to find Chrome again
                $chromePath = Find-ChromeInstallation
                if ($chromePath) {
                    return $chromePath
                } else {
                    Write-Host "Chrome installation completed but executable not found" -ForegroundColor Red
                    return $null
                }
            } else {
                Write-Host "Chrome installation failed with exit code: $($installProcess.ExitCode)" -ForegroundColor Red
                return $null
            }
        } else {
            Write-Host "Failed to download Chrome installer" -ForegroundColor Red
            return $null
        }
        
    } catch {
        Write-Host "Error during Chrome installation: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    } finally {
        # Cleanup
        if (Test-Path $tempDir) {
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Function to find Chrome installation
function Find-ChromeInstallation {
    Write-Host "Looking for Chrome installation..." -ForegroundColor Yellow
    
    $chromePaths = @(
        "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "${env:LOCALAPPDATA}\Google\Chrome\Application\chrome.exe",
        "${env:USERPROFILE}\AppData\Local\Google\Chrome\Application\chrome.exe"
    )
    
    foreach ($path in $chromePaths) {
        if (Test-Path $path) {
            Write-Host "Found Chrome at: $path" -ForegroundColor Green
            return $path
        }
    }
    
    # Try to find via registry
    try {
        $regPath = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe" -ErrorAction SilentlyContinue
        if ($regPath -and (Test-Path $regPath.'(Default)')) {
            Write-Host "Found Chrome via registry at: $($regPath.'(Default)')" -ForegroundColor Green
            return $regPath.'(Default)'
        }
    } catch {
        Write-Host "Registry lookup failed" -ForegroundColor Yellow
    }
    
    Write-Host "Chrome not found in standard locations" -ForegroundColor Red
    return $null
}

# Function to get compatible ChromeDriver version
function Get-CompatibleChromeDriver {
    param([string]$ChromePath)
    
    Write-Host "Getting Chrome version..." -ForegroundColor Yellow
    
    try {
        # Get Chrome version
        $chromeVersion = (Get-ItemProperty $ChromePath).VersionInfo.ProductVersion
        Write-Host "Chrome version: $chromeVersion" -ForegroundColor Cyan
        
        # Extract major version
        $majorVersion = $chromeVersion.Split('.')[0]
        Write-Host "Chrome major version: $majorVersion" -ForegroundColor Cyan
        
        # Download compatible ChromeDriver
        $driverPath = "$env:USERPROFILE\chromedriver_$majorVersion.exe"
        
        if (Test-Path $driverPath) {
            Write-Host "Compatible ChromeDriver already exists: $driverPath" -ForegroundColor Green
            return $driverPath
        }
        
        Write-Host "Downloading ChromeDriver for Chrome $majorVersion..." -ForegroundColor Yellow
        
        # Try to get the latest compatible version
        # For Chrome 115+, use the new Chrome for Testing API
        if ([int]$majorVersion -ge 115) {
            Write-Host "Using Chrome for Testing API for Chrome $majorVersion..." -ForegroundColor Cyan
            
            try {
                # Get available versions from Chrome for Testing API
                $apiUrl = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
                $response = Invoke-RestMethod -Uri $apiUrl -ErrorAction Stop
                
                # Find the latest stable version for this major version
                $compatibleVersions = $response.versions | Where-Object { $_.version -like "$majorVersion.*" }
                
                if ($compatibleVersions.Count -gt 0) {
                    $latestVersion = ($compatibleVersions | Sort-Object version -Descending)[0].version
                    Write-Host "Found compatible ChromeDriver version: $latestVersion" -ForegroundColor Cyan
                    
                    # Find the download URL for Windows chromedriver
                    $chromeDriverDownload = $compatibleVersions | Where-Object { $_.version -eq $latestVersion } | Select-Object -First 1
                    $downloadUrl = $chromeDriverDownload.downloads.chromedriver | Where-Object { $_.platform -eq "win32" } | Select-Object -First 1
                    
                    if ($downloadUrl) {
                        $downloadUrl = $downloadUrl.url
                    } else {
                        throw "No Windows ChromeDriver download found for version $latestVersion"
                    }
                } else {
                    throw "No compatible ChromeDriver versions found for Chrome $majorVersion"
                }
            } catch {
                Write-Host "Chrome for Testing API failed, trying alternative..." -ForegroundColor Yellow
                # Fallback to direct download for latest
                $latestVersion = "$majorVersion.0.0.0"
                $downloadUrl = "https://storage.googleapis.com/chrome-for-testing-public/$latestVersion/win32/chromedriver-win32.zip"
            }
        } else {
            # For older Chrome versions, use the legacy API
            $latestUrl = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_$majorVersion"
            
            try {
                $latestVersion = Invoke-RestMethod -Uri $latestUrl -ErrorAction Stop
                Write-Host "Latest compatible ChromeDriver version: $latestVersion" -ForegroundColor Cyan
            } catch {
                Write-Host "Failed to get latest version, trying generic latest..." -ForegroundColor Yellow
                $latestVersion = Invoke-RestMethod -Uri "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
            }
            
            $downloadUrl = "https://chromedriver.storage.googleapis.com/$latestVersion/chromedriver_win32.zip"
        }
        
        $zipPath = "$env:TEMP\chromedriver_$latestVersion.zip"
        
        Write-Host "Downloading from: $downloadUrl" -ForegroundColor Cyan
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath
        
        # Extract ChromeDriver
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $tempExtractPath = "$env:TEMP\chromedriver_extract_$latestVersion"
        
        if (Test-Path $tempExtractPath) {
            Remove-Item $tempExtractPath -Recurse -Force
        }
        
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempExtractPath)
        
        # Find ChromeDriver executable (might be in a subdirectory for newer versions)
        $extractedDriverPaths = @(
            "$tempExtractPath\chromedriver.exe",
            "$tempExtractPath\chromedriver-win32\chromedriver.exe",
            "$tempExtractPath\chromedriver\chromedriver.exe"
        )
        
        $extractedDriver = $null
        foreach ($path in $extractedDriverPaths) {
            if (Test-Path $path) {
                $extractedDriver = $path
                break
            }
        }
        
        # If not found, search recursively
        if (-not $extractedDriver) {
            $foundDrivers = Get-ChildItem -Path $tempExtractPath -Name "chromedriver.exe" -Recurse -ErrorAction SilentlyContinue
            if ($foundDrivers.Count -gt 0) {
                $extractedDriver = Join-Path $tempExtractPath $foundDrivers[0].FullName
            }
        }
        
        if ($extractedDriver -and (Test-Path $extractedDriver)) {
            Move-Item $extractedDriver $driverPath
            Write-Host "ChromeDriver installed successfully: $driverPath" -ForegroundColor Green
        } else {
            throw "ChromeDriver extraction failed - executable not found in extracted files"
        }
        
        # Cleanup
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        
        return $driverPath
        
    } catch {
        Write-Error "Failed to setup ChromeDriver: $($_.Exception.Message)"
        return $null
    }
}

# Function to setup Selenium with proper error handling
function Initialize-Selenium {
    Write-Host "Initializing Selenium..." -ForegroundColor Yellow
    
    # Check if Selenium module is available
    if (-not (Get-Module -ListAvailable -Name Selenium)) {
        Write-Host "Installing Selenium module..." -ForegroundColor Yellow
        try {
            Install-Module -Name Selenium -Force -Scope CurrentUser -AllowClobber
            Write-Host "Selenium module installed successfully!" -ForegroundColor Green
        } catch {
            Write-Error "Failed to install Selenium module: $($_.Exception.Message)"
            return $false
        }
    } else {
        # Update to latest version
        Write-Host "Updating Selenium module to latest version..." -ForegroundColor Yellow
        try {
            Update-Module -Name Selenium -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Host "Update attempt failed, continuing with existing version..." -ForegroundColor Yellow
        }
    }
    
    try {
        Import-Module Selenium -Force
        Write-Host "Selenium module imported successfully!" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to import Selenium module: $($_.Exception.Message)"
        return $false
    }
}

# Function to extract data using Selenium
function Get-HPDataWithSelenium {
    param(
        [string]$ChromePath,
        [string]$DriverPath
    )

    Write-Host "Starting Selenium browser automation..." -ForegroundColor Yellow

    try {
        # Configure Chrome options
        $chromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions

        # Set Chrome binary location
        $chromeOptions.BinaryLocation = $ChromePath

        # Add arguments
        $chromeOptions.AddArgument("--no-sandbox")
        $chromeOptions.AddArgument("--disable-dev-shm-usage")
        $chromeOptions.AddArgument("--disable-gpu")
        $chromeOptions.AddArgument("--disable-extensions")
        $chromeOptions.AddArgument("--disable-plugins")
        $chromeOptions.AddArgument("--disable-images")
        $chromeOptions.AddArgument("--disable-web-security")
        $chromeOptions.AddArgument("--allow-running-insecure-content")
        $chromeOptions.AddArgument("--disable-features=TranslateUI")
        $chromeOptions.AddArgument("--disable-ipc-flooding-protection")
        $chromeOptions.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36")

        if ($HeadlessBrowser) {
            $chromeOptions.AddArgument("--headless")
            Write-Host "Running in headless mode..." -ForegroundColor Cyan
        } else {
            Write-Host "Running in visible mode..." -ForegroundColor Cyan
        }

        # Create ChromeDriver service
        $driverDirectory = [System.IO.Path]::GetDirectoryName($DriverPath)
        $driverFileName = [System.IO.Path]::GetFileName($DriverPath)

        # Create service with explicit driver path
        $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($driverDirectory, $driverFileName)
        $service.HideCommandPromptWindow = $true

        # Set environment variables to help with driver detection
        $env:CHROMEDRIVER_PATH = $DriverPath

        # Start browser with explicit service
        Write-Host "Starting Chrome browser with ChromeDriver: $DriverPath" -ForegroundColor Yellow
        try {
            $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $chromeOptions)
        } catch {
            Write-Host "Failed to start with service, trying direct path..." -ForegroundColor Yellow
            # Fallback: try setting driver path in environment
            $env:PATH = "$driverDirectory;$env:PATH"
            $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($chromeOptions)
        }

        # Set timeouts
        $driver.Manage().Timeouts().ImplicitWait = [TimeSpan]::FromSeconds(10)
        $driver.Manage().Timeouts().PageLoad = [TimeSpan]::FromSeconds(30)

        try {
            # Navigate to the page
            Write-Host "Navigating to: $url" -ForegroundColor Cyan
            $driver.Navigate().GoToUrl($url)

            # Wait for page to load and JS to execute
            Write-Host "Waiting for page to load and JavaScript to execute (${WaitTimeSeconds}s)..." -ForegroundColor Cyan
            Start-Sleep -Seconds $WaitTimeSeconds

            # Save page source for debugging
            $pageSource = $driver.PageSource
            $pageSource | Out-File -FilePath "selenium_page_source.html" -Encoding UTF8
            Write-Host "Page source saved to: selenium_page_source.html" -ForegroundColor Cyan

        } finally {
            # Clean up
            Write-Host "Closing browser..." -ForegroundColor Yellow
            $driver.Quit()
        }

    } catch {
        Write-Error "Selenium automation failed: $($_.Exception.Message)"
        if ($ShowProgress) {
            Write-Host "Stack trace:" -ForegroundColor Red
            Write-Host $_.Exception.StackTrace
        }
        return @()
    }
}

# Function to extract HP compatibility data from HTML and output JSON only
function Extract-HPWin11CompatibilityFromHtml {
    param(
        [string]$HtmlFile = "selenium_page_source.html",
        [string]$JsonOutput = "HP_Win11_Compatibility.json"
    )
    Write-Host "Extracting HP Windows 11 compatibility data from HTML..." -ForegroundColor Yellow
    if (-not (Test-Path $HtmlFile)) {
        Write-Host "Error: HTML file '$HtmlFile' not found!" -ForegroundColor Red
        return
    }
    $htmlContent = Get-Content -Path $HtmlFile -Raw
    Add-Type -AssemblyName System.Web
    $compatibilityData = @()
    $headerPattern = '<th class="entry[^"]*"\s+id="([^"]+)"\s+[^>]*>\s*<p class="p">([^<]+)</p>\s*</th>'
    $headerMatches = [regex]::Matches($htmlContent, $headerPattern)
    $headers = @()
    foreach ($match in $headerMatches) {
        $headerId = $match.Groups[1].Value
        $headerText = $match.Groups[2].Value.Trim()
        $headers += @{ Id = $headerId; Text = $headerText }
    }
    if ($headers.Count -eq 0) {
        Write-Host "No table headers found!" -ForegroundColor Red
        return
    }
    $rowPattern = '<tr class="row">\s*((?:<td[^>]*>.*?</td>\s*)+)\s*</tr>'
    $rowMatches = [regex]::Matches($htmlContent, $rowPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    foreach ($rowMatch in $rowMatches) {
        $rowHtml = $rowMatch.Groups[1].Value
        $cellPattern = '<td class="entry[^"]*"\s+headers="([^"]+)"\s+[^>]*>\s*<p class="p">([^<]*)</p>\s*</td>'
        $cellMatches = [regex]::Matches($rowHtml, $cellPattern)
        if ($cellMatches.Count -eq $headers.Count) {
            $rowData = [ordered]@{}
            for ($i = 0; $i -lt $cellMatches.Count; $i++) {
                $headerId = $cellMatches[$i].Groups[1].Value
                $cellValue = $cellMatches[$i].Groups[2].Value.Trim()
                $header = $headers | Where-Object { $_.Id -eq $headerId }
                if ($header) { $rowData[$header.Text] = $cellValue }
            }
            if ($rowData.Count -gt 0) { $compatibilityData += $rowData }
        }
    }
    $jsonOutputObj = @{
        ExtractedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        SourceUrl = "https://support.hp.com/us-en/document/ish_4890350-4890415-16"
        TableTitle = "Business PCs and workstations tested with Windows 11"
        Headers = $headers | ForEach-Object { $_.Text }
        DataCount = $compatibilityData.Count
        Data = $compatibilityData
    }
    $jsonString = $jsonOutputObj | ConvertTo-Json -Depth 10 -Compress:$false
    $jsonString | Out-File -FilePath $JsonOutput -Encoding UTF8
    Write-Host "JSON output: $JsonOutput" -ForegroundColor Green
}

# Main execution
try {
    # Initialize Selenium
    if (-not (Initialize-Selenium)) {
        throw "Failed to initialize Selenium"
    }
    
    # Find Chrome installation
    $chromePath = Find-ChromeInstallation
    if (-not $chromePath) {
        Write-Host "Chrome not found in standard locations. Attempting to install..." -ForegroundColor Yellow
        $chromePath = Install-GoogleChrome
        
        if (-not $chromePath) {
            throw "Failed to install Chrome browser. Please install Google Chrome manually."
        }
    }
    
    # Get compatible ChromeDriver
    $chromeDriverPath = Get-CompatibleChromeDriver -ChromePath $chromePath
    if (-not $chromeDriverPath) {
        throw "Failed to setup ChromeDriver"
    }
    
    # Extract data using Selenium
    Write-Host "Starting automated data extraction..." -ForegroundColor Yellow
    $null = Get-HPDataWithSelenium -ChromePath $chromePath -DriverPath $chromeDriverPath
    
    # After saving selenium_page_source.html, call the extraction function
    Extract-HPWin11CompatibilityFromHtml -HtmlFile "selenium_page_source.html" -JsonOutput "HP_Win11_Compatibility.json"
} catch {
    Write-Error "Automation failed: $($_.Exception.Message)"
    Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
    Write-Host "1. Make sure Google Chrome is installed" -ForegroundColor Cyan
    Write-Host "2. Check your internet connection" -ForegroundColor Cyan
    Write-Host "3. Try running with -HeadlessBrowser `$false to see what's happening" -ForegroundColor Cyan
    Write-Host "4. Increase wait time with -WaitTimeSeconds parameter" -ForegroundColor Cyan
    Write-Host "5. Run as administrator if there are permission issues" -ForegroundColor Cyan
    Write-Host "6. Try running with -ShowProgress for detailed output" -ForegroundColor Cyan
}

Write-Host "`nScript completed." -ForegroundColor Green