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
        $volumeFolder = $mangeFolder + "\Volume " + ($chapter.Name).Substring(4,2)
        Write-Host "Volume Folder: $volumeFolder "
        Write-Host " "

        # check if exist and create chapter folder
        $chapterFolder = $volumeFolder + "\Chapter " +  ($chapter.Name).Substring(10,($chapter.Name.IndexOf("(") - 11))
        Write-Host "Chapter Folder: $chapterFolder "
        Write-Host " "

        if(-Not (Test-Path -Path $chapterFolder)) {
			Write-Host "Creating Chapter Folder $chapterFolder"
			Write-Host " "
			New-Item -ItemType Directory -Path $chapterFolder -Force
		}

        # get the prefix for the renaming archive Vol.02 Ch.0010  Page.19.png
		$namePrefix = ($chapter.Name).Substring(0,($chapter.Name.IndexOf("-") - 1)) + " Page."
		
		# process every file in the chapter folder 
		Write-Host "Copy files to $chapterFolder"
		foreach ($file in (Get-ChildItem -literalpath $chapter.FullName)) {	
			# copy the file to the chapter folder with the new name
			$newName = $namePrefix + $file
            $filePath = $chapterFolder + "\" + $newName
            Write-Host "File Path: $filePath"
			Write-Host "$file copy as $newName"
            Copy-Item -LiteralPath $file.FullName -Destination $filePath
		}
        Write-Host " "
	}
}

# after finished the processing, compressing to cbz 
foreach ($mangaFolder in (Get-ChildItem -Path $processingFolder)){
    if(-Not (Test-Path -Path $finishedFolder)){ 
        New-Item -ItemType Directory -Path $finishedFolder -Force 
    }
    Write-Host "Compressing $mangaFolder"
    $destinationPath = $finishedFolder + "\" + $mangaFolder + ".zip"
    Compress-Archive -Path $mangaFolder.FullName -DestinationPath $destinationPath
    Rename-Item -Path $destinationPath -NewName ($mangaFolder.Name + ".cbz")
}

# delete temporary folder
Remove-Item -Path $processingFolder -Recurse

# call kcc app to convert the cbz to mobi file format 
# TODO - this must be configurable, not only in data, but also in execution
# foreach ($manga in (Get-ChildItem -Path $finishedFolder)){
#     Write-Host "Converting $manga from cbz to mobi file format"
#     .\kcc-c2e_5.6.3.exe $manga.FullName --profile K11 --manga-style --hq --stretch --cropping 0 --title $manga.Name --format MOBI 
# }