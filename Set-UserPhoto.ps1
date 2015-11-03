#Script to export O365 users in a csv file formated for MigrationWiz

#----------------------------------------------------------
#STATIC VARIABLES
#----------------------------------------------------------
#Output level 3 = Verbose, 2 = Warning, 1 = Error
$outputlevel = 3 

$scriptpath = "C:\Scripts"
$scriptname = "Set-UserPhoto"

$PhotosPath = "C:\UsersPhoto"
$csvfilename = "Users.csv"
if ($MSOLCred -eq $null)
{
$MSOLCred = get-credential
}
if ($Exsession -ne $null)
{
$ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxymethod=rps -Credential $MSOLCred -Authentication Basic -AllowRedirection
}

#----------------------------------------------------------
#AUTO CALCULATED VARIABLES
#----------------------------------------------------------
$hostname = hostname
$x=$y=$z=0

$csvpath = "$scriptpath\$csvfilename"

#----------------------------------------------------------
#START FUNCTIONS
#----------------------------------------------------------

Function write-log
{
param ($log, $type)
	if ($type -eq 1)
	{
		$typename = "WARNING"
		if ($outputlevel -ge 2)
		{
			write-host $log -foreground yellow
		}
	}
	elseif ($type -eq 2)
	{
		$typename = "ERROR"
		if ($outputlevel -ge 1)
		{
			write-host $log -foreground red
		}
	}
	else
	{
		$typename = "INFORMATION"
		if ($outputlevel -ge 3)
		{
			write-host $log
		}
	}
	$date = get-date
	$filename = get-date -uformat %m-%Y
	$filename = "$scriptpath\$scriptname-" + "$filename" + ".log"
	$log = "$date" + "     " + "$typename" + "     " +"$log"
	$log >> $filename
}

#Starting Script
write-log "----------------------Starting Script $scriptname -------------------------------"

#Starting Powershell Remote session with Exchange Online
try
{
#import-PSSession $ExSession
}
Catch
{
Write-Log "Error connecting to Exchange Online" 1
Write-Log $error[0]
Exit
}

#Check the repository of photos
try
{
$photos = get-ChildItem $PhotosPath
}
Catch
{
Write-Log "Error when looking for photos" 1
Write-Log $error[0]
}

foreach ($photo in $photos)
    {
    $user=$null
    $fullname = $photo.name.substring(0,($photo.name.length-4))
    $fullname

    $user = get-user $fullname -ErrorAction SilentlyContinue
    


    if ($user -ne $null)
    {

        $userpic = get-userphoto $user.UserPrincipalName -ErrorAction SilentlyContinue

         if(($user -ne $null) -and ($userpic -eq $null))
        {
        write-log "User recognised $fullname"
        try{
        Set-userPhoto $user.UserPrincipalName -PictureData ([System.IO.File]::ReadAllBytes($photo.fullname)) -confirm:$false
        write-log "Picture set $fullname"
        }
        catch{
        write-log "Error setting the photo"
        write-log $error[0]
        }

        $y++
        }
        elseif(($user -ne $null) -and ($userpic -ne $null))
        {
        write-log "User already has a photo $fullname"
        $z++
        }
       
    }
    else
    {
    write-log "Can't find user $fullname" 1
    mv $Photo.fullname "$PhotosPath\nouser"
    }
$x++
   
}

write-log "Find $x photos files and $y users recognised $z users already have a photo" 
write-log "---------------------------END OF SCRIPT---------------------------"

#write-log "Can't find user $fullname"
#$LastName = $fullname.substring($fullname.indexof(" "), ($fullname.length - $fullname.indexof(" ")))
#$LastName
#$user = get-user -filter "Name -Like '*$LastName*'"#
#$ans = Read-Host "Do you want to use Photo $fullname for user $user"