param (
    [Parameter(Mandatory=$true)] 
    [string] $srcPath = '.',
    [string] $dstPath = '.\finished',
    [switch] $byManga = $false,
    [switch] $executeKcc = $false,
    [switch] $askMetadata = $false,
    [switch] $askMetadataByVolume = $false,
    [string] $srcCsvMetadata = $null,
    [switch] $ignoreNotFoundMetadataCsv = $false
)

# Constants
$TittleKey = "Tittle" 
$SeriesKey = "Series"
$WritterKey = "Writter"
$PencillerKey = "Penciller"
$InkerKey = "Inker"
$ColoristKey = "Colorist"
$SummaryKey = "Summary"
$processingFolder = $srcPath + "\processing"
$CSVData = $null

# Private Functions
function Write-ComicInfo {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String] $Path,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Tittle,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String] $Series,
        [String] $Volume,
        [String] $Number,
        [String] $Writer,
        [String] $Penciller,
        [String] $Inker,
        [String] $Colorist,
        [string] $Summary
    )

    # declaring the document to create for the metadata
    $Path = $Path + "\ComicInfo.xml"
    $xmlWriter = New-Object System.XMl.XmlTextWriter($Path, $Null)

    # Configuring the format 
    $xmlWriter.Formatting = 'Indented'
    $xmlWriter.Indentation = 1
    $XmlWriter.IndentChar = "`t"

    $xmlWriter.WriteStartDocument()

    # Write the metadata of the comic / manga
    $xmlWriter.WriteStartElement("ComicInfo") # Root element

    # Attributes for the root element
    $xmlWriter.WriteAttributeString("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
    $xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

    # write only the given data
    $xmlWriter.WriteElementString("Tittle", $Tittle)
    $xmlWriter.WriteElementString("Series", $Series)
    if(-Not ([string]::IsNullOrEmpty($Volume))){
        $xmlWriter.WriteElementString("Volume", $Volume)
    }
    if(-Not ([string]::IsNullOrEmpty($Number))){
        $xmlWriter.WriteElementString("Number", $Number)
    }
    if(-Not ([string]::IsNullOrEmpty($Writer))){
        $xmlWriter.WriteElementString("Writer", $Writer)
    }
    if(-Not ([string]::IsNullOrEmpty($Penciller))){
        $xmlWriter.WriteElementString("Penciller", $Penciller)
    }
    if(-Not ([string]::IsNullOrEmpty($Inker))){
        $xmlWriter.WriteElementString("Inker", $Inker)
    }
    if(-Not ([string]::IsNullOrEmpty($Colorist))){
        $xmlWriter.WriteElementString("Colorist", $Colorist)
    }
    if(-Not ([string]::IsNullOrEmpty($Summary))){
        $xmlWriter.WriteElementString("Summary", $Summary)
    }

    # Finish and save the xml file
    $xmlWriter.WriteEndElement()
    $xmlWriter.WriteEndDocument()
    $xmlWriter.Flush()
    $xmlWriter.Close()
}
function Read-ComicInfo {
    param (
        [string] $Tittle,
        [string] $Series,
        [String] $Writer,
        [String] $Penciller,
        [String] $Inker,
        [String] $Colorist,
        [String] $Summary
    )

    Write-Host "Input the Manga / Comic Information"

    $TittleDefault = $Tittle
    if(!($Tittle = Read-Host -Prompt "Tittle: [$Tittle]")) { $Tittle = $TittleDefault }
    $SeriesDefault = $Series
    if (!($Series = Read-Host -Prompt "Series: [$Series]")) { $Series = $SeriesDefault }
    $WriterDefault = $Writer
    if (!($Writer = Read-Host -Prompt "Writer: [$Writer]")) { $Writer = $WriterDefault }
    $PencillerDefault = $Penciller
    if (!($Penciller = Read-Host -Prompt "Penciller: [$Penciller]")) { $Penciller = $PencillerDefault }
    $InkerDefault = $Inker
    if (!($Inker = Read-Host -Prompt "Inker: [$Inker]")) { $Inker = $InkerDefault }
    $ColoristDefault = $Colorist
    if (!($Colorist = Read-Host -Prompt "Colorist: [$Colorist]")) { $Colorist = $ColoristDefault }
    $SummaryDefault = $Summary
    $SummaryShortText = $Summary
    if($Summary.Length -gt 15){ $SummaryShortText = $Summary.Substring(0, 15) }
    if (!($Summary = Read-Host -Prompt "Summary: [$SummaryShortText]")) { $Summary = $SummaryDefault }

    [hashtable] $metadata = @{$TittleKey=$Tittle; $SeriesKey=$Series; $WritterKey=$Writer; $PencillerKey=$Penciller; $InkerKey=$Inker; $ColoristKey=$Colorist; $SummaryKey=$Summary}
    return $metadata
}
function Read-ComicInfo-CSV {
    param (
        [String] $Series,
        [String] $Volume
    )

    Write-Host "Looking info in the CSV for $Series Volume $Volume"

    # Load the CSV file to get the info
    if(!$CSVData) {
        $CSVData = Import-Csv -Path $srcCsvMetadata -Delimiter ";"
    }

    # Search for Series and volume
    $Results = $null
    if($Volume -and $Series){
        $Results = $CSVData | Where-Object {($_.Series -eq $Series) -and ($_.Volume -eq $Volume)}
    }
    # Check if there is any result for the manga
    if(!$Results) {
        # in case of no result, search for only the series
        $Results = $CSVData | Where-Object {($_.Series -eq $Series)}
    }
    # Check if there is any result for the manga 
    if(!$Results -and !$ignoreNotFoundMetadataCsv){
        # ask to manually input the manga / comic metadata
        return Read-ComicInfo $Series $Series
    }

    # return the info recovered
    $Tittle = $Results[0].Tittle
    if(!$Tittle){
        if(!$Volume){
            $Tittle = $Results[0].Series
        }
        else {
            $Volume = [String]$Volume
            while ($Volume.Length -le 1) { $Volume = "0" + $Volume }
            $Tittle = $Results[0].Series + " Volume " + $Volume
        }
    }
    [hashtable] $metadata = @{$TittleKey=$Tittle; $SeriesKey=$Results[0].Series; $WritterKey=$Results[0].Writer; $PencillerKey=$Results[0].Penciller; $InkerKey=$Results[0].Inker; $ColoristKey=$Results[0].Colorist; $SummaryKey=$Results[0].Summary}
    return $metadata

}

# main script
Write-Host "Manga Prepare Script"
Write-Host "Author: Adrian Jesus"
Write-Host "script to preparer the chapters folders of a manga, downloaded by HakuNeko Software, into volumes guided by the name of the folders to track the volumes"

Write-Host " "

Write-Host "starting processing of the manga folder: $srcPath"

# check / create destination directory
if(-Not (Test-Path -Path $dstPath)){
    New-Item -ItemType Directory -Path $dstPath -Force
}
Write-Host "destinantion Folder: $dstPath"

if(Test-Path -Path $processingFolder){
    Remove-Item $processingFolder -Recurse
}

# list all the folder in location
foreach ($folder in (Get-ChildItem -LiteralPath $srcPath -Directory)) {
    Write-Host "Processing $folder"
	Write-Host " "

    # creating the processing folder path for the manga
    $mangeFolder = $processingFolder + "\" + $folder.Name
    Write-Host "Manga Folder: $mangeFolder"
    Write-Host " "

    foreach ($chapter in (Get-ChildItem $folder.FullName -Directory)) {
        Write-Host "Processing $chapter"
		Write-Host " "
		
		# check if exist and create volume folder
        $addVolumeNumber = $false
        $volumeFolder = ""
        if ($chapter.Name.Contains("Vol")) {
            $volumeFolder = $mangeFolder + "\Volume " + ($chapter.Name).Substring(4,2)
        }
        else {
            $volumeFolder = $mangeFolder + "\Volume 01"
            $addVolumeNumber = $true
        }
        
        Write-Host "Volume Folder: $volumeFolder "
        Write-Host " "
        
        # check if exist and create chapter folder
        $chapterFolder = ""
        if ($chapter.Name.Contains("(")) {
            if ($chapter.Name.Contains("Vol")) {
                $chapterFolder = $volumeFolder + "\Chapter " +  ($chapter.Name).Substring(10,($chapter.Name.IndexOf("(") - 11))
            }
            else {
                $chapterFolder = $volumeFolder + "\Chapter " +  ($chapter.Name).Substring(3,($chapter.Name.IndexOf("(") - 4))
            }
        }
        else{
            $chapterFolder = $volumeFolder + "\Chapter " +  ($chapter.Name).Substring(10)
        }
        
        # Clean the trailing spaces and dots
        while (($chapterFolder -match '\.$') -or ($chapterFolder -match '\s$')) {
            $chapterFolder = $chapterFolder.Substring(0, $chapterFolder.Length - 1)
        }

        Write-Host "Chapter Folder: $chapterFolder "
        Write-Host " "

        if(-Not (Test-Path -Path $chapterFolder)) {
			Write-Host "Creating Chapter Folder $chapterFolder"
			Write-Host " "
			New-Item -ItemType Directory -Path $chapterFolder -Force
            Write-Host " "
		}

        # get the prefix for the renaming archive Vol.02 Ch.0010  Page.19.png
		$namePrefix = $chapter.Name
        if ($chapter.Name.Contains("-")){
            $namePrefix = ($chapter.Name).Substring(0,($chapter.Name.IndexOf("-") - 1)) + " Page."
        }
        elseif ($chapter.Name.Contains("(")) {
            $namePrefix = ($chapter.Name).Substring(0,($chapter.Name.IndexOf("(") - 1)) + " Page."
        }
        if($addVolumeNumber){
            $namePrefix = "Vol.01 " + $namePrefix
        }
        Write-Host "Name Prefix: $namePrefix"
        Write-Host " "
		
		# process every file in the chapter folder 
		Write-Host "Copy files to $chapterFolder"
        $pageNumber = 1
		foreach ($file in (Get-ChildItem -literalpath $chapter.FullName)) {	
			# copy the file to the chapter folder with the new name
            $strPageNumber = $pageNumber.ToString()
            if($strPageNumber.Length -eq 1){
                $strPageNumber = "0" + $strPageNumber
            }
			$newName = $namePrefix + $strPageNumber + $file.Extension
            $filePath = $chapterFolder + "\" + $newName
            Write-Host "File Path: $filePath"
			Write-Host "$file copy as $newName"
            Copy-Item -LiteralPath $file.FullName -Destination $filePath
            $pageNumber += 1
		}
        Write-Host " "
    }
}

# after finished the processing, compressing to cbz 
foreach ($mangaFolder in (Get-ChildItem -Path $processingFolder)){
    if(-Not (Test-Path -Path $dstPath)){ 
        New-Item -ItemType Directory -Path $dstPath -Force 
    }

    [hashtable] $metadata = @{}
    if($askMetadata -and !$askMetadataByVolume){
        $metadata = Read-ComicInfo $mangaFolder.Name $mangaFolder.Name
    }

    # check if is going to compress by manga folder or by volume folder
    if($byManga) {
        Write-Host "Compressing $mangaFolder"
        $destinationPath = $dstPath + "\" + $mangaFolder + ".zip"

        if(Test-Path -Path ($dstPath + "\" + $mangaFolder.Name + ".cbz")){
            Remove-Item -Path ($dstPath + "\" + $mangaFolder.Name + ".cbz")
        }

        if($srcCsvMetadata){
            $metadata = Read-ComicInfo-CSV $mangaFolder.Name 
        }

        # write the metadata / comic info of the volume folder
        if(-Not ($metadata.Count -eq 0)) {
            Write-ComicInfo $mangaFolder.FullName $metadata[$TittleKey] $metadata[$SeriesKey] $null $null $metadata[$WritterKey] $metadata[$PencillerKey] $metadata[$InkerKey] $metadata[$ColoristKey] $metadata[$SummaryKey]
        }

        # Compress the manga / comic
        Compress-Archive -Path $mangaFolder.FullName -DestinationPath $destinationPath
        Rename-Item -Path $destinationPath -NewName ($mangaFolder.Name + ".cbz")
    }
    else {
        $destinationFolder = $dstPath + "\" + $mangaFolder.Name
        New-Item -ItemType Directory -Path $destinationFolder -Force
        foreach ($volumeFolder in (Get-ChildItem -Path $mangaFolder.FullName)){

            # Check if there is the need to ask for input metadata manually
            if($askMetadataByVolume) {
                if ($metadata.Count -eq 0) {
                    $metadata = Read-ComicInfo $mangaFolder.Name $mangaFolder.Name
                }
                else {
                    $metadata = Read-ComicInfo $metadata[$TittleKey] $metadata[$SeriesKey] $metadata[$WritterKey] $metadata[$PencillerKey] $metadata[$InkerKey] $metadata[$ColoristKey] $metadata[$SummaryKey]
                }
            }

            $volumeNumber = [int]$volumeFolder.Name.Substring(7)
            if($srcCsvMetadata){
                $metadata = Read-ComicInfo-CSV $mangaFolder.Name $volumeNumber
            }

            Write-Host "Compressing $mangaFolder $volumeFolder"
            $destinationPath = $destinationFolder + "\" + $mangaFolder.Name + " " + $volumeFolder.Name
            $destinationZipFile = $destinationPath + ".zip"
            $destinationCbzFileName = $mangaFolder.Name + " " + $volumeFolder.Name + ".cbz"

            if (Test-Path -Path ($destinationFolder + "\" + $destinationCbzFileName)){
                Remove-Item -LiteralPath ($destinationFolder + "\" + $destinationCbzFileName)
            }

            # write the metadata / comic info of the volume folder
            if(-Not ($metadata.Count -eq 0)){
                Write-ComicInfo $volumeFolder.FullName $metadata[$TittleKey] $metadata[$SeriesKey] $volumeNumber $null $metadata[$WritterKey] $metadata[$PencillerKey] $metadata[$InkerKey] $metadata[$ColoristKey] $metadata[$SummaryKey]
            }

            Compress-Archive -Path $volumeFolder.FullName -DestinationPath $destinationZipFile
            Rename-Item -Path $destinationZipFile -NewName ($destinationCbzFileName)
        }
    }
}

# delete temporary folder
Remove-Item -Path $processingFolder -Recurse

# call kcc app to convert the cbz to mobi file format if asked for it
if ($executeKcc) {
    if($byManga) {
        foreach ($manga in (Get-ChildItem -Path $dstPath)){
            $mobiFile = $manga.FullName.replace('.cbz', '.mobi')
            if(Test-Path -Path $mobiFile){
                Remove-Item -Path $mobiFile
            }
            Write-Host "Converting $manga from cbz to mobi file format"
            .\kcc-c2e_5.6.3.exe $manga.FullName --profile K11 --manga-style --stretch --cropping 0 --title $manga.Name --format MOBI
        }
    }
    else {
        foreach ($manga in (Get-ChildItem -Path $dstPath)) {
            foreach ($volume in (Get-ChildItem -Path $manga.FullName)){
                $mobiFile = $volume.FullName.replace('.cbz', '.mobi')
                if(Test-Path -Path $mobiFile){
                    Remove-Item -Path $mobiFile
                }
                Write-Host "Converting $volume from cbz to mobi file format"
                .\kcc-c2e_5.6.3.exe $volume.FullName --profile K11 --manga-style --stretch --cropping 0 --title $volume.Name --format MOBI
            }
        }
    }
}
