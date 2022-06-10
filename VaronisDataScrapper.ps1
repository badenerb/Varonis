#######################################################################
#                                                                     #
#                      Varonis Data Scrapper                          #
#          Make sure to change the deskpath to your desktop path      #
#    This program will scrape data from the CSVs that Varonis gives   #
#                                                                     #
#######################################################################

#Variable Declerations

$checkWordlist = false
$platform = 3

#####################################################
$DESKPATH = "C:\Users\" + $env:USERNAME + "\OneDrive - Vermeer Corporation\Desktop"#
#CHANGE THIS if it does not work^^^^#################

#Functions

#DONE
#Startup function that creates the startup menu and graphic.
function startup
{
    #Graphic/Intro
    clear
    Write-Host @"

                 Welcome to the Varonis Data Scrapper
 
                       Press Enter to begin                                                                                                                                                                                                                                             
                                                                                                                                                                
     @@@&                                                                       
     @@@@@#                                                                     
      @@@@@@/                                                                   
       /@@@@@@,                                                                 
         @@@@@@@                                                                
           #@@@@@@                                                              
             %@@@@@@                                                            
               #@@@@@#                                                          
                .@@@@@@/                                                        
                  ,@@@@@@,                                                      
                    *@@@@@@.                                                    
                      @@@@@@@####################                               
 %((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((@     
 /                                                                        &     
 / @@/    .@@*  @@@@@@@@  *@@@@@@@# @@@@@@@@@@/ @@@@   .@@*.@@,,@@@@@@@#  &     
 / @@/    .@@*.@@/.  ./@@,*@@.  @@# @@#,  .,@@/ @@@@@%,.@@*.@@,,@@*.      &     
 / @@/   *#@@*.@@@@@@@@@@,*@@#%@@@, @@(     @@/ @@/ /@@@@@*.@@,./@@@@@#*  &     
 / @@#**@@@#  .@@*    ,@@,*@@@@@*   @@%*  .*@@/ @@/   /#@@*.@@,    .*@@#  &     
 /  &@@@&*    .&@*    ,@&.,&@.*&@@# &@@@@@@@@&* &@/    .@&*.@&,,&@@@@@&(  &     
 /                                                                        &     
 ,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,   

                                                                 Created by:
                                                                   Baden Erb
"@
    Read-Host
}

#DONE
#Function that will scrape human users from a Varonis CSV file
function userScrape
{
    clear
    $inputFile = Read-Host -Prompt "What is the name of the Varonis csv output file (excluding the file extension)"
    $FILENAME = "$inputFile.csv"
    cd ~
    Set-Location $DESKPATH

    if(!(Test-Path -Path ".\$FILENAME"))
    {
        Write-Host "Ensure the Varonis output file, $FILENAME, is on your desktop or exists"
        read-host
        return
    }

    Write-Host "Processing the stale users. Please be patient"

    $date = Get-Date -Format "_MM_dd_yyyy_HH_mm_ss"
    $newFile = "$inputFile$date.csv"
    New-Item -name "$newFile" | Out-Null
    Write-Host "Created new file for the data called $newFile"

    Write-output ("Domain,Name,User,Email,Department,Manager") | Out-File ".\$newFile" -encoding ascii

    foreach($line in Get-Content .\$FILENAME)
    {
        $splitArray = @()
        $splitArray =  $line.Split(",")
        $username = $splitArray[2]
        $email = $splitArray[3]
        if(($username -match "[A-Z]{2,3}[0-9]{3,6}") -and ($email.length -lt 50) -and ($username.length -lt 8))
        {
            $input = ""
            $domain = $splitArray[0]
            $name = $splitArray[1]
            $department = $splitArray[4]
            $manager = $splitArray[5].Trim("vermeermfg.com\")
            $input = "$domain,$name,$username,$email,$department,$manager,$input"
            Write-output ($input) | Out-File ".\$newFile" -encoding ascii -Append
        }
    }

    write-host "`r`nProccessing complete. Press Enter to return to the menu"
    read-host
}

#BROKEN
#Function will eventually find open Sharepoint sites from the Varonis CSV file
function shareScrape
{
    clear
    $inputFile = Read-Host -Prompt "What is the name of the Varonis csv output file (excluding the file extension)"

    $FILENAME = "$inputFile.csv"
    cd ~
    Set-Location $DESKPATH

    if(!(Test-Path -Path ".\$FILENAME"))
    {
        Write-Host "Ensure the Varonis output file, $FILENAME, is on your desktop or exists"
        read-host
        return
    }

    while($true)
    {
        clear
        Write-Host "Select an option to filter your data, when you are done press 3"
        Write-Host "`r`n1. Change the platform to filter by"
        Write-Host "2. Change if you want to check files against a dangerous filename wordlist"
        Write-Host "3. Start the processing of the file."
        $opt = Read-Host -Prompt "Option #"
        
        if($opt -eq 1)
        {
            Write-Host "What platform would you like to filter the data by?"
            Write-Host "1 for OneDrive, 2 for SharePoint, or 3 for both"
            $platform = Read-Host -Prompt "Selection #"
        }
        elseif($opt -eq 2)
        {
            Write-Host "Would you like to check file names against a wordlist? Y for Yes or N for No"
            $wordlistCheck = read-host -Prompt "Y or N"
            if(($wordlistCheck -eq "Y") -or ($wordlistCheck -eq "y"))
            {
                $checkWordlist = true
            }
        }
        elseif($opt -eq 3)
        {
            Write-Host "Press enter to begin processing"
            Read-Host
            break
        }
        else
        {
            Write-Host "Invalid selection, press enter to try again."
            Read-host
            continue
        }
    }

    clear
    Write-Host "Processing the shared data. Please be patient"

    $date = Get-Date -Format "_MM_dd_yyyy_HH_mm_ss"
    $newFile = "$inputFile$date.csv"
    New-Item -name "$newFile" | Out-Null
    Write-Host "Created new file for the data called $newFile"
    Write-output ("Platform,Event Type,Object Type,Event Operation,Object Name (Event On),Path,Event Description,User Name (Event By),Event Time,Is Alerted,Blacklisted Location,Alert Severity,Threat Model Name") | Out-File ".\$newFile" -encoding ascii
    Write-Host "Creating an array"
    $filterArray=@()
    foreach($line in Get-Content .\$FILENAME)
    {
        if ($line.ReadCount -eq 1)
        {
            continue
        }
        $split = $line.Split(",")
        $plat = $split[0].Trim()
        if($platform -eq "1")
        {
            if($plat.Equals("OneDrive"))
            {
                $filterArray += $line
            }
            else
            {
                continue
            }
        }
        elseif($platform -eq "2")
        {
            if($plat.Equals("SharePoint Online"))
            {
                $filterArray += $line
            }
            else
            {
                continue
            }
        }
        else
        {
            $filterArray += $line
        }
    }
    Write-Host "Array created succesfully"
    if($wordlistCheck)
    {
        Write-Host "Checking files for potentially dangerous names"
        Write-Host "Downloading the wordlist."
        Invoke-WebRequest -Uri https://raw.githubusercontent.com/Bo0oM/fuzz.txt/master/fuzz.txt -OutFile .\fuzz.txt
        Write-Host "scanning the object and path names for potentially dangerous names"
        $badFiles = @()
        foreach ($j in $filterArray)
        {
            $splitJ = $j.Split(",")
            $object = $splitJ[4]
            $path = $splitJ[5]
            foreach ($line in Get-Content ./fuzz.txt)
            {
                if(($object.Equals($line)) -or ($path.Equals($line)))
                {
                    $badFiles += $j
                }
            }
            
        }
        Remove-Item "./fuzz.txt"
        Write-host "Good objects and paths filtered out of array."
        Write-Host "Putting array into $newfile"
        foreach($i in $badFiles)
        {
            Write-output ("$i") | Out-File ".\$newFile" -encoding ascii -Append
        }

    }
    else
    {
        Write-Host "Putting array into $newfile"
        foreach($i in $filterArray)
        {
            Write-output ("$i") | Out-File ".\$newFile" -encoding ascii -Append
        }
    }
    Write-Host "Processing complete. Press Enter to return to the menu."
    Read-Host
}

#DONE
function sharePointLook
{
    clear
    $inputFile = Read-Host -Prompt "What is the name of the Varonis CSV output file (excluding the file extension)"
    $FILENAME = "$inputFile.csv"
    cd ~
    Set-Location $DESKPATH

    if(!(Test-Path -Path ".\$FILENAME"))
    {
        Write-Host "Ensure the Varonis output file, $FILENAME, is on your desktop or exists"
        read-host
        return
    }

    $index = 2
    foreach ($line in Get-Content .\$FILENAME)
    {
        #Get the URL
        $url = getURL -fp .\$FILENAME -ln $index

        #Test the URL
        try
        {
            $WebResponse = Invoke-WebRequest $url
        }
        catch
        {
            $index++
            continue
        }
        Write-Host $url
        $content = $WebResponse.Content
        if (!($content -like '*<title>Sign in to your account</title>*'))
        {
            write-host "This $url is open"
        }
        $index++
    }
    Write-Host "Processing complete, press enter to continue."
    Read-Host
}

#Done
#Returns the full URL in a given file on the given line number
function getURL
{
    param($fp, $ln)
    $indexOfServer = 0
    $indexOfPath = 0
    $header = Get-Content $fp | Select -First 1
    $urlLine = Get-Content $fp | Select -First $ln | select -Last 1
    $splitHead = $header.Split(",")
    for($i = 0; $i -lt $splitHead.Length; $i++)
    {
        $j = $splitHead[$i]
        if($j -eq "File Server")
        {
            $indexOfServer = $i
        }
        elseif($j -eq "Path")
        {
            $indexOfPath = $i
        }
    }
    $splitLine = $urlLine.Split(",")
    $fullURL = $splitLine[$indexOfServer] + $splitLine[$indexOfPath]
    $fullURL = $fullURL.replace('\','/')
    return $fullURL
}

#Menu
clear
startup
clear
while ($true)
{
    clear
    Write-Host @"
  __  __       _          
 |  \/  | __ _(_)_ __     
 | |\/| |/ _` | | '_ \    
 | |  | | (_| | | | | |   
 |_|  |_|\__,_|_|_| |_|   
 |  \/  | ___ _ __  _   _ 
 | |\/| |/ _ \ '_ \| | | |
 | |  | |  __/ | | | |_| |
 |_|  |_|\___|_| |_|\__,_|                  

Select an option from below
        
      1. Scrape User data to get human users
      2. Scrape shared data
      3. Find truley open access SharePoint sites
      4. Exit

"@

    $sel = Read-Host -Prompt "Option #"

    if($sel -eq 1)
    {
        userScrape
    }
    elseif($sel -eq 2)
    {
        shareScrape
    }
    elseif($sel -eq 3)
    {
        sharePointLook
    }
    elseif($sel -eq 4)
    {
        break
    }
    else
    {
        Write-Host "Invalid selection, press enter to try again."
        read-host
        continue
    }
}
