###	Cerakote-Fetcher
###	v1.0
###	8/31/2023
###
### 	Written by Dale Stevens
###
###	This Powershell script will download all the entries of colors for h-series cerakote as well as the image swatches. 

###	Optional Global Variables: 
###	allcats = [ cera_h_series, cera_c_series, cera_elite, cera_v_series ]
###	orderby = [ created_at, rank ]
###	orderdir = [ asc, desc ]
### 	hide_discounted = [ true, false ]
### 	fields = [ id, url, featured_image_url, name, sku, categories.**, discontinued, quantity_available, base_quantity.**, product_id, swatch_product.id, swatch_product.enabled, swatch_product.hero_only, pricing.** ]


# Define global variables

# Base URL for the JSON data
$global:baseUrl = "https://www.cerakote.com/api/proxify/domains/cerakote/shop/products"

# Query parameters
$global:queryParams = @{
    "allcats" = "cera_h_series"
    "orderby" = "rank"
    "orderdir" = "desc"
    "hide_discontinued" = "false"
    "fields" = "sku, name, featured_image_url"
}

# Maximum number of pages to fetch
$global:maxPages = 10

# Function to check if SKU exists in the CSV
function SKU-Exists {
    param (
        [string]$sku,
        [string]$csvPath
    )

    try {
        $csvData = Import-Csv -Path $csvPath -Encoding UTF8
        return ($csvData | Where-Object { $_.sku -eq $sku }).Count -gt 0
    } catch {
        return $false
    }
}

# Function to sanitize the 'name' field by removing special characters
function Sanitize-Name {
    param (
        [string]$name
    )

    # Remove special characters, including non-alphanumeric characters
    $sanitizedName = [System.Text.RegularExpressions.Regex]::Replace($name, '[^a-zA-Z0-9\s]', '')

    return $sanitizedName
}

# Function to download images and count new entries
function Download-Images {
    param (
        [string]$csvPath,
        [string]$outputFolder
    )

    # Create the output folder if it doesn't exist
    if (-not (Test-Path -Path $outputFolder -PathType Container)) {
        New-Item -Path $outputFolder -ItemType Directory
    }

    # Initialize counters for new entries and downloaded images
    $newEntryCount = 0
    $downloadedImageCount = 0

    try {
        # Import the CSV data with UTF-8 encoding
        $csvData = Import-Csv -Path $csvPath -Encoding UTF8

        # Iterate through each row in the CSV
        foreach ($row in $csvData) {
            $sku = $row.sku
            $name = $row.name
            $imageUrl = $row.featured_image_url

            # Construct the full path to save the image
            $imagePath = Join-Path -Path $outputFolder -ChildPath "$sku - $name.jpg"

            # Check if the image file already exists
            if (-not (Test-Path -Path $imagePath)) {
                # Download the image from the URL
                Invoke-WebRequest -Uri $imageUrl -OutFile $imagePath

                Write-Host "Downloaded image: $sku - $name.jpg"

                # Increment the counter for downloaded images
                $downloadedImageCount++
            } else {
                Write-Host "Image already exists: $sku - $name.jpg"
            }
        }

        # Display appropriate messages based on downloaded images
        if ($downloadedImageCount -eq 0) {
            Write-Host "No new images were downloaded."
        } else {
            Write-Host "Downloaded $downloadedImageCount new images to folder: $outputFolder"
        }
    } catch {
        Write-Host "Error: $_"
    }
}

# Get the current directory as the base path
$basePath = (Get-Location).Path

# Define the path for the CSV file where you want to save the data
$csvPath = Join-Path -Path $basePath -ChildPath "cerakote_export_cleaned.csv"

# Create an array to store custom objects for CSV export
$customObjects = @()

# Loop through pages
for ($page = 1; $page -le $maxPages; $page++) {
    # Modify the query parameters to include the current page
    $queryParams.page = $page

    # Create a UriBuilder object and add query parameters
    $uriBuilder = [System.UriBuilder]::new($baseUrl)
    $queryString = [System.Uri]::UnescapeDataString($uriBuilder.Query)

    foreach ($param in $queryParams.GetEnumerator()) {
        if ($queryString -ne "") {
            $queryString += "&"
        }

        $queryString += [System.Uri]::EscapeDataString($param.Key) + "=" + [System.Uri]::EscapeDataString($param.Value)
    }

    $uriBuilder.Query = $queryString

    try {
        # Fetch JSON data from the URL for the current page
        $jsonResponse = Invoke-RestMethod -Uri $uriBuilder.Uri -Method Get -ContentType "application/json"

        # Check if the response contains data
        if ($jsonResponse -and $jsonResponse.data -is [System.Array]) {
            foreach ($item in $jsonResponse.data) {
                $sku = $item.sku
                $name = (Sanitize-Name -name $item.name)  # Apply the Sanitize-Name function

                # Check if SKU exists in the CSV file
                if (-not (SKU-Exists -sku $sku -csvPath $csvPath)) {
                    # Create a custom object with selected properties
                    $customObject = [PSCustomObject]@{
                        "sku"             = $sku
                        "name"            = $name
                        "featured_image_url" = $item.featured_image_url
                    }

                    $customObjects += $customObject
                }
            }
        } else {
            Write-Host "Page ${page}: No valid data was retrieved from the API."
        }
    } catch {
        Write-Host "Page ${page}: Error: $_"
    }
}

# Export the custom objects to a CSV file with sanitized 'name' and UTF-8 encoding
if ($customObjects.Count -gt 0) {
    $customObjects | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "JSON data successfully converted, sanitized, and saved to CSV: $csvPath"
} else {
    Write-Host "No new entries were added to the CSV."
}

# Specify the full path for the output folder
$outputFolder = Join-Path -Path $basePath -ChildPath "swatch_images"

# Call the function to download images
Download-Images -csvPath $csvPath -outputFolder $outputFolder
