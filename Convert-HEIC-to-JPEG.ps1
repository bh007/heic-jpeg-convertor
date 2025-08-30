# Define the source folder where your HEIC files are located.
# You can change this path to any folder you need.
$sourceFolder = "C:\Path\To\Your\Photos"

# Get all HEIC files in the specified folder.
# The -Recurse parameter can be added if you want to include subfolders.
$heicFiles = Get-ChildItem -Path $sourceFolder -Filter "*.heic"

# If no HEIC files are found, show a message and exit.
if ($null -eq $heicFiles) {
    Write-Host "No HEIC files found in '$sourceFolder'."
    return
}

Write-Host "Found $($heicFiles.Count) HEIC file(s) to convert."

# Create WIC factory object.
$WICFactory = New-Object -ComObject WIC.ImagingFactory

foreach ($file in $heicFiles) {
    # Create the full path for the new JPEG file.
    # The .jpg extension replaces the .heic extension.
    $jpegPath = $file.DirectoryName + "\" + $file.BaseName + ".jpg"

    try {
        # Create a decoder for the HEIC file.
        $decoder = $WICFactory.CreateDecoderFromFilename($file.FullName, $null, 
            [System.IO.FileAccess]::Read, [WIC.DecodeOptions]::WICDecodeMetadataCacheOnDemand)
        
        # Get the first frame of the image.
        $frame = $decoder.GetFrame(0)

        # Create an encoder for the JPEG format.
        $encoder = $WICFactory.CreateEncoder([WIC.ContainerFormat]::WICContainerFormatJpeg, $null)
        
        # Create a file stream to write the output JPEG.
        $jpegStream = New-Object -ComObject ADODB.Stream
        $jpegStream.Open()
        $encoder.Initialize($jpegStream)

        # Create a new frame and copy the pixel data from the original frame.
        $newFrame = $encoder.CreateNewFrame($null)
        $newFrame.Initialize()
        $newFrame.SetSize($frame.GetSize().Width, $frame.GetSize().Height)
        $newFrame.WriteSource($frame, $null)

        # Commit changes and save the file.
        $newFrame.Commit()
        $encoder.Commit()

        # Save the stream to the specified path.
        $jpegStream.SaveToFile($jpegPath, [ADODB.SaveOptionsEnum]::adSaveCreateOverWrite)
        $jpegStream.Close()

        Write-Host "✅ Converted '$($file.Name)' to '$($file.BaseName).jpg'"

    } catch {
        # If conversion fails, show an error message.
        Write-Error "Failed to convert '$($file.Name)'. Error: $($_.Exception.Message)"
    }
}

Write-Host "`nConversion process complete."
