

function Import-Excel   # import the excel and stores everything in a array called $filestore.
{
  param (
    [string]$FileName,   # need these params filled out
    [string]$WorksheetName,
    [bool]$DisplayProgress = $true
  )

  if ($FileName -eq "") { #check for filename
    throw "Please provide path to the Excel file"
    Exit
  }

  if (-not (Test-Path $FileName)) { #check if file exists
    throw "Path '$FileName' does not exist."
    exit
  }

  $FileName = Resolve-Path $FileName #Get path to filename
  $excel = New-Object -com "Excel.Application" # new object created for excel.
  $excel.Visible = $false
  $workbook = $excel.workbooks.open($FileName) # open workbook

  if (-not $WorksheetName) { # check worksheet name if no name given use the first sheet
    Write-Warning "Defaulting to the first worksheet in workbook."
    $sheet = $workbook.ActiveSheet
  } else {
    $sheet = $workbook.Sheets.Item($WorksheetName)
  }
  
  if (-not $sheet) # if no sheet throw error and exit 
  {
    throw "Unable to open worksheet $WorksheetName"
    exit
  }
  #variables set
  $sheetName = $sheet.Name
  $columns = $sheet.UsedRange.Columns.Count
  $lines = $sheet.UsedRange.Rows.Count
  
  Write-Warning "Worksheet $sheetName contains $columns columns and $lines lines of data" # write to sreen to show how many lines we are going through 
  
  $fields = @() # initate array to be added into
  
  for ($column = 1; $column -le $columns; $column ++) { # recrusively go through every column
    $fieldName = $sheet.Cells.Item.Invoke(1, $column).Value2
    if ($fieldName -eq $null) {
      $fieldName = "Column" + $column.ToString()
    }
    $fields += $fieldName # add to array for column names
  }
  
  $line = 2 # start at line 2 for data
  
  
  for ($line = 2; $line -le $lines; $line ++) { # starting at line 2 for data 
    $values = New-Object object[] $columns # create a object for the data to be inputted
    for ($column = 1; $column -le $columns; $column++) { # go through each column for data
      $values[$column - 1] = $sheet.Cells.Item.Invoke($line, $column).Value2 #pull data
    }  
  
    $row = New-Object psobject # each row is object
    $fields | foreach-object -begin {$i = 0} -process {
      $row | Add-Member -MemberType noteproperty -Name $fields[$i] -Value $values[$i]; $i++ # add row of data to each object
    }
    $row # display data
    $percents = [math]::round((($line/$lines) * 100), 0) # show progress 
    if ($DisplayProgress) {
      Write-Progress -Activity:"Importing from Excel file $FileName" -Status:"Imported $line of total $lines lines ($percents%)" -PercentComplete:$percents
    }
  }
  $workbook.Close()
  $excel.Quit()
}

function FindFiles 
{ # Pull GUIDS from FileStore array and then search directories for FileGUIDS. Return GUID and Directory location then copy and zip and rename to original name which exists in filestore
    
    $UpdatedArray = @() # inintalize new array to be made from $filestore
    $item = New-Object PSObject
    for ($line = 0; $line -le $filestore.Count;$line++)
    { # go trough each line in $filestore
        
       $UpdatedArray += New-Object psobject -Property @{ # creating new object array of strictly file guid, file name, and date
            FileGuid = $filestore[$line].FileGuid
            FileName = $filestore[$line].FileName
            Date = $filestore[$line].FileCreated 
       }
    }
    echo $UpdatedArray
    
    for ($line =0;$line -le $UpdatedArray.Count; $line++)
    {   #pull date search for file guid in respective folders
        $CopyCount = 1
        $fileGuid = $UpdatedArray[$line].FileGuid # pulll the guid
        $fileDate = $UpdatedArray[$line].Date # pull the date
        $fileName = $UpdatedArray[$line].FileName # pull the name
        Write-Host " -----------------------------------------------" -ForegroundColor Yellow
        echo $fileDate #checks
        echo ""
        echo $fileGuid #checks
        echo ""
        echo $fileName #checks
       
        $count = 1
        $filePath = $FileStorePath+'.'+$fileDate
        $file =  Get-ChildItem -Recurse -Force $filePath -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and  ( $_.Name -like "*$fileGuid*") }
        if($file -eq $null)
        {
            $filePath =  $FileStorePath
            echo "Coudln't find file trying $filePath"
            $file =  Get-ChildItem -Recurse -Force $filePath -ErrorAction SilentlyContinue | Where-Object { ($_.PSIsContainer -eq $false) -and  ( $_.Name -like "*$fileGuid*") }
            if($file -eq $null){
                $lineSplit = "------------------------------------------------------------"
                echo $fileDate,$fileGuid,$fileName,$lineSplit | out-file -Append -filepath $MissingFileList

            } 
			else {echo "Error: Ran into problem with looking for file: $fileDate,$fileGuid,$fileName" >> $MissingFileList }
        }
        else
        {
            echo "Locating File in $filePath"
            Get-ChildItem -Recurse -Force $filePath | Where-Object { ($_.PSIsContainer -eq $false) -and  ( $_.Name -like "*$fileGuid*") } | ForEach-Object{
                $NewPath = $destination # make new path the destination
                echo $NewPath
                $DirectoryCheck = Get-ChildItem $NewPath #pull all directories in path
                echo $fileDate
                $NewDirectory = $NewPath+"\"+$fileDate # New Directory Name
                echo $NewDirectory
                if($DirectoryCheck.Name -ccontains $fileDate) #check if Directory exists
                {
                    $DirectoryCheck2 = Get-ChildItem $NewDirectory # check directory for file names 
                    Copy-Item -Path $_.FullName -Destination $NewDirectory # copy  item to destination.
                    if($DirectoryCheck2.Name -ccontains $fileName)
                    {
                        Rename-Item $NewDirectory\$fileGuid".dat" $fileName"Copy"$copyCount # rename to original file name
                        $copyCount++
                    }
                    else
                    {
                         Rename-Item $NewDirectory\$fileGuid".dat" $fileName # rename to original file name
                    }
                }
                else # if directory does not exist
                {
                    New-Item $NewDirectory -ItemType directory # create new directory
                    $DirectoryCheck2 = Get-ChildItem $NewDirectory # check directoru for file names 
                    Copy-Item -Path $_.FullName -Destination $NewDirectory # copy  item to destination.
                    if($DirectoryCheck2.Name -ccontains $fileName)
                    {
                        Rename-Item $NewDirectory\$fileGuid".dat" $fileName"Copy"$copyCount # rename to original file name
                        $copyCount++
                    }
                    else
                    {
                         Rename-Item $NewDirectory\$fileGuid".dat" $fileName # rename to original file name
                    }
                   
               }
            }
            
        }
    }
}

function Zipping{
    Get-ChildItem -Path $destination | ForEach-Object{
        $Zipped = $NameOfZip
        $source = $destination
        $DirectoryZipped = $destination+"\"+$zipped
        Add-Type -assembly "system.io.compression.filesystem"
        [io.compression.ZipFile]::CreateFromDirectory($source, $DirectoryZipped)
    }
}       

function execute {
    FindFiles
	Zipping
}


$NameOfZip = "Images"
$MissingFileList = 'C:\nullFiles.txt'
$FileStorePath = 

$destination = 'C:\test'
$filestore = Import-Excel 'C:\TESTING.csv'
execute
