param (
    [switch]$byVolume = $false,
    [switch]$executeKcc = $false
)

Write-Host "Manga Prepare Script"
Write-Host "Author: Adrian Jesus"
Write-Host "script to preparer the chapters folders of a manga, downloaded by HakuNeko Software, into volumes guided by the name of the folders to track the volumes"
Write-Host "THIS MUST BE RUN IN THE ROOT FOLDER WHERE ARE ALL THE MANGA"

Write-Host " "

Write-Host "starting processing of the manga folder"

$processingFolder = "ZZ-processing"
$finishedFolder = "ZZ-finished"
if(Test-Path -Path $processingFolder){
    Remove-Item -Path $processingFolder -Recurse
}
if(Test-Path -Path $finishedFolder){
    Remove-Item -Path $finishedFolder -Recurse
}

# list all the folder in location
foreach ($folder in (Get-ChildItem -Directory)) {
    Write-Host "Processing $folder"
	Write-Host " "

    # creating the processing folder path for the manga
    $mangeFolder = $processingFolder + "\" + $folder.Name
    Write-Host "Manga Folder: $mangeFolder"
    Write-Host " "

    foreach ($chapter in (Get-ChildItem $folder -Directory)) {
        Write-Host "Processing $chapter"
		Write-Host " "
		
		# check if exist and create volume folder
        $volumeFolder = ""
        if ($chapter.Name.Contains("Vol")) {
            $volumeFolder = $mangeFolder + "\Volume " + ($chapter.Name).Substring(4,2)
        }
        else {
            $volumeFolder = $mangeFolder
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
        
        Write-Host "Chapter Folder: $chapterFolder "
        Write-Host " "

        # TODO: handle the case where the folder contais a "." and other characters
        # the Copy-Item can't handle that type of names for some reason
        # if the chapter name is only dots is no problem aparently 
        # TODO: delete the last white space if any
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
    if(-Not (Test-Path -Path $finishedFolder)){ 
        New-Item -ItemType Directory -Path $finishedFolder -Force 
    }

    # check if is going to compress by manga folder or by volume folder
    if($byVolume){
        $destinationFolder = $finishedFolder + "\" + $mangaFolder.Name
        New-Item -ItemType Directory -Path $destinationFolder -Force
        foreach ($volumeFolder in (Get-ChildItem -Path $mangaFolder.FullName)){
            Write-Host "Compressing $mangaFolder $volumeFolder"
            $destinationPath = $destinationFolder + "\" + $mangaFolder.Name + " " + $volumeFolder.Name
            $destinationZipFile = $destinationPath + ".zip"
            $destinationCbzFileName = $mangaFolder.Name + " " + $volumeFolder.Name + ".cbz"
            Compress-Archive -Path $volumeFolder.FullName -DestinationPath $destinationZipFile
            Rename-Item -Path $destinationZipFile -NewName ($destinationCbzFileName)
        }
    }
    else {
        Write-Host "Compressing $mangaFolder"
        $destinationPath = $finishedFolder + "\" + $mangaFolder + ".zip"
        Compress-Archive -Path $mangaFolder.FullName -DestinationPath $destinationPath
        Rename-Item -Path $destinationPath -NewName ($mangaFolder.Name + ".cbz")
    }
}

# delete temporary folder
Remove-Item -Path $processingFolder -Recurse

# call kcc app to convert the cbz to mobi file format if asked for it
if ($executeKcc) {
    if($byVolume) {
        foreach ($manga in (Get-ChildItem -Path $finishedFolder)) {
            foreach ($volume in (Get-ChildItem -Path $manga.FullName)){
                Write-Host "Converting $volume from cbz to mobi file format"
                .\kcc-c2e_5.6.3.exe $volume.FullName --profile K11 --manga-style --stretch --cropping 0 --title $volume.Name --format MOBI
            }
        }
    }
    else{
        foreach ($manga in (Get-ChildItem -Path $finishedFolder)){
            Write-Host "Converting $manga from cbz to mobi file format"
            .\kcc-c2e_5.6.3.exe $manga.FullName --profile K11 --manga-style --stretch --cropping 0 --title $manga.Name --format MOBI
        }
    }
}
