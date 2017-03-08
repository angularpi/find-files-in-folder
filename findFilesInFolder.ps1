#filename contains the name ("BaseName") of the files I need to find and extract from originalFolder
$fileList = "C:\Users\angularpi\list.csv"

#original folder contains the 1,000~ files
$origFolder = "C:\Users\angularpi\origFolder\"
#newFolder is initially empty
$newFolder = "C:\Users\angularpi\newFolder\"

#check if newFolder exists, and create it if it does not
$test = Test-Path $newFolder
if (!$test){
    md $newFolder
}

#import list into variable
$list = Import-Csv $fileList

#main loop
foreach ($item in $list){
    $item = $item.BaseName #BaseName is the name of the column in the csv file
    
    #compose the item's file path
    $file = $origFolder + $item + ".xl*" #the * ignores the end of the extension in case there are .xlsx or .xlsm files
   
    #check is file exists in origFolder
    $test = Test-Path($file)

    #if file exists then copy it to newFolder
    if ($test = $TRUE){
        cp $file $newFolder

        #show results in screen
        Write-Host($item + "| " + $test)
    }
}
