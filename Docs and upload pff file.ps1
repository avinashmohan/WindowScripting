####################Doc functions####################


Function OpenWordDoc($Filename)

{

$Word=NEW-Object –comobject Word.Application


Return $Word.documents.open($Filename)


}




Function SearchAWord($Document,$findtext,$replacewithtext)

{ 

  $FindReplace=$Document.ActiveWindow.Selection.Find

  $matchCase = $false;

  $matchWholeWord = $true;

  $matchWildCards = $false;

  $matchSoundsLike = $false;

  $matchAllWordForms = $false;

  $forward = $true;

  $format = $false;

  $matchKashida = $false;

  $matchDiacritics = $false;

  $matchAlefHamza = $false;

  $matchControl = $false;

  $read_only = $false;

  $visible = $true;

  $replace = 2;

  $wrap = 1;

  $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl)

}



Function SaveAsWordDoc($Document,$FileName)

{

$Document.Saveas([REF]$Filename)

$Document.close()
taskkill.exe /F /IM WINWORD.exe

}





####################Doc function ends################


function BLfile{


try
{

###################

cd "C:\Program Files (x86)\WinSCP"
   Add-Type -Path "WinSCPnet.dll"

    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = "XXX.XX.X.XX"
        UserName = "test"
        Password = "test"
        #port="22"
        SshHostKeyFingerprint = "ssh-rsa 2048 a8:bc:ac:c4:68:5c:8b:64:bd:a4:d1:36:4b:28:a1:c1"
        
    }
$session = New-Object WinSCP.Session

$session.Open($sessionOptions)




########################



$localadd=$localaddold.Text

cd "$localadd\"

logMessage "Please don't '/' ahead of the local location."

##BL_CTR_PROD_file_editing
$versioning=$textboxVersionLiv.Text
$versionnum=$versioning.Trim("v")
Write-Host "Version number:$versionnum"

$Date1= Get-date -format dd/MM/yyyy
Write-Host "Date1:$Date1"

$Date2=Get-date -format yyyyMMdd
Write-Host "Date1:$Date2"

$hour=Get-date -format hhmm
Write-Host "Hour:$hour"
#$hour=$hourformat.trim("AM")

##################

#$locallocation=$localaddold.Text
$locallocation=$rootDir


    if(Test-Path $locallocation)
    {
    Write-Host "Loacal address exist"
    logMessage "Loacal address exist"

    $file1="XXX/YYY/ZZZ/location of file1.docx"
    $file2="XXX/YYY/ZZZ/location of file2.docx"
    $localPath_out=$rootDir
     $transferResult3 = $session.GetFiles($file1, $localPath_out , $False, $transferOptions)

        $transferResult4 = $session.GetFiles($file2, $localPath_out , $False, $transferOptions)

    if ($transferResult3.Error -eq $Null){
                       
                        foreach ($transfer3 in $transferResult3.Transfers)
                            {
                                Write-Host ("Download of {0} succeeded." -f $transfer3.FileName)
                                logMessage ("Download of {0} succeeded." -f $transfer3.FileName)  
                            }
                            }
                            else
                            {
                             Write-Host "Download of BL-PROD failed"
                            }


                            
    if ($transferResult4.Error -eq $Null){
                       
                        foreach ($transfer4 in $transferResult4.Transfers)
                            {
                                Write-Host ("Download of {0} succeeded." -f $transfer4.FileName)
                                logMessage ("Download of {0} succeeded." -f $transfer4.FileName)  
                            }
                            }
                            ELSE
                            {
                            Write-Host "Download of BL-G9 failed"
                            }


$VersionName="VERSION-"+$textboxVersionLiv.Text+"-"+$currentDate
$remote1 = "/XXX/YYY/ZZZ/$VersionName/Testloc"
Write-host "Version number si : $VersionName"
    $wildcard = "*.zip"
     $x=2
$files = $session.EnumerateRemoteFiles($remote1, $wildcard, [WinSCP.EnumerationOptions]::None)

        # Any file matched?
        
        if ($files.Count -gt 0)
        {
       $localadd=$rootDir
        taskkill.exe /F /IM WINWORD.exe

         $localaddress1=$localadd+"Test_file1.docx"
        Write-Host "Loacl address is:$localaddress"

         $Doc=OpenWordDoc -Filename $localaddress1
         #$Doc.Font.Color = "WDColorDarkBlue"
       
                      $Table = $Doc.Tables.item(7)

        foreach ($fileInfo in $files)
            {
                   $name=$fileInfo.name
                    Write-Host "file name:$name"

                    if ($name -like "ETR_TOOLS-*")
                    {
                    $stringTools=$name
                   
                      $Table.Cell($x,1).Range.Text = $name
                    $x=$x+1
                    Write-Host "Loop number:$x"


                     
                ##  $zipremove=$name.trim(".zip")
                ##  Write-Host "Zip removed string is $zipremove"
                ##
                ##  $indexlast= $zipremove.LastIndexOf("-")
                ##  Write-Host "Last Index:$indexlast"
                ##
                ##  $length=$zipremove.length
                ##  Write-Host "Last Index:$length"
                ##
                ##  $finalstring=$zipremove.substring($indexlast)
                ##
                ##  Write-HOst "Time hour:$finalstring"
                ##
                ##  
                ##
                ##
                ##   $finalToolstring=$zipremove.substring(0,$indexlast)
                ##
                ##   Write-HOst "finaltoolstring:$finalToolstring"
                ##
                ##   
                ##  $indexlst2= $finalToolstring.LastIndexOf("-")
                ##   Write-HOst "Index2 :$indexlst2"
                ##
                ##   $datefinal= $finalToolstring.substring($indexlst2+1)
                ##   Write-HOst "Datefinal= $datefinal"
                ##
                ##   $a=$datefinal
                ##   $Date1=[datetime]::ParseExact($a, "dd/MM/yy", $null)
                           ##     Write-Host "Date1:$Date1"



                    }

                    if ($name -like "table_name_row1-*")
                    {
                    $stringBDD=$name
                    
                      $Table.Cell($x,1).Range.Text = $name
                      $x=$x+1
                      Write-Host "Loop number:$x"

                    }

                    if ($name -like "table_name_row2-*")
                    {
                   $stringAPP=$name
                    
                      $Table.Cell($x,1).Range.Text = $name
                      $x=$x+1
                      Write-Host "Loop number:$x"


                    }

                    if ($name -like "table_name_row3*")
                    {
                    $stringWEB=$name
                    
                      $Table.Cell($x,1).Range.Text = $name
                       $x=$x+1
                       Write-Host "Loop number:$x"


                    }
                  
                   if ($name -like "table_name_row4*")
                    {
                    $stringADP=$name
                    
                      $Table.Cell($x,1).Range.Text = $name
                      $x=$x+1
                      Write-Host "Loop number:$x"


                    }
                  
                   
          


                  

            }
SearchAWord –Document $Doc -findtext "x.xx.x" -replacewithtext $versionnum
SearchAWord –Document $Doc -findtext "DD/MM/YYYY" -replacewithtext $Date1
SearchAWord –Document $Doc -findtext "yyyymmdd" -replacewithtext $Date2
SearchAWord –Document $Doc -findtext "hhmm" -replacewithtext $finalstring
$Savename1="BL_CTR_"+$versionnum+"_PROD.docx"
write-host "Saving the file as $Savename1 "
SaveAsWordDoc –document $Doc –Filename "$localPath_out\$Savename1"
logMessage "Created $Savename1 sucessfully"
            }


##################

$x2=2
if ($files.Count -gt 0)
        {
       
        taskkill.exe /F /IM WINWORD.exe
        $localaddress=$localadd+"BL_CTR_x_xx_x_PPD_G9.docx"
        Write-Host "Loacl address is:$localaddress"

         $Doc2=OpenWordDoc -Filename $localaddress
        
                      $Table2 = $Doc2.Tables.item(7)

        foreach ($fileInfo in $files)
            {
                   $name=$fileInfo.name
                    Write-Host "file name:$name"

                    if ($name -like "ETR_TOOLS-*")
                    {
                    $stringTools=$name
                   
                      $Table2.Cell($x2,1).Range.Text = $name
                    $x2=$x2+1
                    Write-Host "Loop number:$x2"


                     
                ##  $zipremove=$name.trim(".zip")
                ##  Write-Host "Zip removed string is $zipremove"
                ##
                ##  $indexlast= $zipremove.LastIndexOf("-")
                ##  Write-Host "Last Index:$indexlast"
                ##
                ##  $length=$zipremove.length
                ##  Write-Host "Last Index:$length"
                ##
                ##  $finalstring=$zipremove.substring($indexlast)
                ##
                ##  Write-HOst "Time hour:$finalstring"
                ##
                ##  
                ##
                ##
                ##   $finalToolstring=$zipremove.substring(0,$indexlast)
                ##
                ##   Write-HOst "finaltoolstring:$finalToolstring"
                ##
                ##   
                ##  $indexlst2= $finalToolstring.LastIndexOf("-")
                ##   Write-HOst "Index2 :$indexlst2"
                ##
                ##   $datefinal= $finalToolstring.substring($indexlst2+1)
                ##   Write-HOst "Datefinal= $datefinal"
                ##
                ##   $a=$datefinal
                ##   $Date1=[datetime]::ParseExact($a, "dd/MM/yy", $null)
                           ##     Write-Host "Date1:$Date1"



                    }

                    if ($name -like "table_name_row1-*")
                    {
                    $stringBDD=$name
                    
                      $Table2.Cell($x2,1).Range.Text = $name
                      $x2=$x2+1
                      Write-Host "Loop number:$x2"

                    }

                    if ($name -like "table_name_row2-*")
                    {
                    $stringAPP=$name
                    
                      $Table2.Cell($x2,1).Range.Text = $name
                      $x2=$x2+1
                      Write-Host "Loop number:$x2"


                    }

                    if ($name -like "table_name_row3*")
                    {
                    $stringWEB=$name
                    
                      $Table2.Cell($x2,1).Range.Text = $name
                       $x2=$x2+1
                       Write-Host "Loop number:$x2"


                    }
                  
                   if ($name -like "table_name_row4*")
                    {
                    $stringADP=$name
                    
                      $Table2.Cell($x2,1).Range.Text = $name
                      $x2=$x2+1
                      Write-Host "Loop number:$x2"


                    }
                  
                   
          


                  

            }
        SearchAWord –Document $Doc2 -findtext "x.xx.x" -replacewithtext $versionnum
SearchAWord –Document $Doc2 -findtext "DD/MM/YYYY" -replacewithtext $Date1
SearchAWord –Document $Doc2 -findtext "yyyymmdd" -replacewithtext $Date2
SearchAWord –Document $Doc2 -findtext "hhmm" -replacewithtext $finalstring
$Savename2="BL_CTR_"+$versionnum+"_PPD_G9.docx"
SaveAsWordDoc –document $Doc2 –Filename "$localPath_out\$Savename2"
logMessage "Created $Savename2 sucessfully"
            }













$ETRremotepath="/XXX/YYY/ZZZ/"

        # Get list of files in the directory
        $directoryInfo = $session.ListDirectory($ETRremotepath) 
 
        # Select the most recent file
        
      
        
        $latest =
            $directoryInfo.Files |
            Where-Object {   $_.Name -match "VERSION-" } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -first 1

        # Any file at all?
        if ($latest -eq $Null)
        {
            Write-Host "No file found"
            logMessage "No file found"
            exit 1
        }

        # Download the selected file
         Write-Host "File found:$latest"
         $filepath="/XXX/YYY/ZZZ/$latest/10-BL/"
         $localPath_out1="$localadd\$Savename1"
         $localPath_out2="$localadd$Savename2"
         
         if ($session.FileExists($filepath))
        {
        Write-Host "file path exist for $latest"
        logMessage "file path exist for $latest"

        #######Upload package files###########

        # Upload files
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

        $transferResult1 = $session.PutFiles($localPath_out1, $filepath, $False, $transferOptions)
           Write-Host "file path :$Savename1 uploaded sucessfully to $filepath"
           logMessage "file path :$Savename1 uploaded sucessfully to $filepath"

           $transferResult2 = $session.PutFiles($localPath_out2, $filepath, $False, $transferOptions)
           Write-Host "file path :$Savename2 uploaded sucessfully to $filepath"
           logMessage "file path :$Savename2 uploaded sucessfully to $filepath"
        # Throw on any error
        $transferResult1.Check()

        # Print results
        foreach ($transfer1 in $transferResult1.Transfers)
        {
           Write-Host ("Upload of {0} succeeded" -f $transfer1.FileName)
            #logMessage ("Upload of {0} succeeded" -f $transfer.FileName)
        }


        $transferResult2.Check()

        # Print results
        foreach ($transfer2 in $transferResult2.Transfers)
        {
           Write-Host ("Upload of {0} succeeded" -f $transfer2.FileName)
            #logMessage ("Upload of {0} succeeded" -f $transfer.FileName)
        }





       #######Upload package files Ends###########

        }
        else 
        {
        Write-Host "file path does not exists,will create new directory"
        logMessage "file path does not exists,will create new directory"

        #BLfile directory creation

        Write-Host ("File create old folder")
            logMessage "Creating 10-BL Directory."
            $createpathold= "/XXX/YYY/ZZZ/$latest/10-BL/"
            #if((test-path $createpathold) -eq $false) { $session.CreateDirectory($createpathold)  }
            
            #New-Item -Path $createpathold -force -ItemType "Directory"

            $session.CreateDirectory($createpathold)

        
         #######Upload package files###########

        # Upload files
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

        $transferResult1 = $session.PutFiles($localPath_out1, $filepath, $False, $transferOptions)
           Write-Host "file path :$Savename1 uploaded sucessfully to $filepath"
           logMessage "file path :$Savename1 uploaded sucessfully to $filepath"

           $transferResult2 = $session.PutFiles($localPath_out2, $filepath, $False, $transferOptions)
           Write-Host "file path :$Savename2 uploaded sucessfully to $filepath"
           logMessage "file path :$Savename2 uploaded sucessfully to $filepath"
        # Throw on any error
        $transferResult1.Check()

        # Print results
        foreach ($transfer1 in $transferResult1.Transfers)
        {
           Write-Host ("Upload of {0} succeeded" -f $transfer1.FileName)
            logMessage ("Upload of {0} succeeded" -f $transfer1.FileName)
        }


           $transferResult2.Check()

        # Print results
        foreach ($transfer1 in $transferResult2.Transfers)
       {
           Write-Host ("Upload of {0} succeeded" -f $transfer2.FileName)
            logMessage ("Upload of {0} succeeded" -f $transfer2.FileName)
        }

       #######Upload package files Ends###########
            }


}
    else 
    {
    Write-Host "Loacal address doesnot exist"
    logMessage "Loacal address doesnot exist"
    [System.Windows.Forms.MessageBox]::Show("ERROR:Please check the local address,where the package is present.
e.g.D:\Users\user.name\Desktop\Uploadfile") 
   
    }



}

catch
{
Write-Host ("Error: {0}" -f $_.Exception.Message)
   

}

finally 
{
#####

}




}

