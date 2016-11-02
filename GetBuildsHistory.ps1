<#
--------------------------------------------------------------------------------- 
Project: BuildsHistory                                               
Version: 1.0                                                         
FileId:  f:\Delme\CleanBCD\BuildsHistory                             
When:    October 10, 2016,  Monday,  09:28, Windows build 14942                             
Who:     Oleg Kulikov, sysprg@live.ru                                                                                          
--------------------------------------------------------------------------------- 
This is PShell version of the BuildsHistory.js which was finished at
October 7, 2016, some hours before Build 14942 became available at WU servers.
--------------------------------------------------------------------------------- 
10.0.14393.321
[MAJOR VERSION].[MINOR VERSION].[BUILD NUMBER].[REVISION NUMBER]
---------------------------------------------------------------------------------
#>

<#
---------------------------------------------------------------------------------
Function converts DateTime stamps in in Registry subkeys like
"Source OS (Updated on 11/7/2015 18:23:38)" into "07112015_182338" 
PShell date format strings: "ddMMyyyy_HHmmss" 
---------------------------------------------------------------------------------
#>
function UTCStringToStandard( [string]$utc )
{ 
   $tmp = $utc.split( "/" )
   if ( $tmp[ 0 ].Length -eq 1 ){ $tmp[ 0 ] = "0" + $tmp[ 0 ] }
   if ( $tmp[ 1 ].Length -eq 1 ){ $tmp[ 1 ] = "0" + $tmp[ 1 ] }
   $tmp = $tmp[ 1 ] + $tmp[ 0 ] + $tmp[ 2 ]
   $tmp = $tmp.split( " " )
   $tmp = $tmp[ 0 ] + "_" + $tmp[ 1 ] 
   $tmp = $tmp.split( ":" ) 
   $tmp = $tmp[ 0 ] + $tmp[ 1 ] + $tmp[ 2 ]  
   return $tmp
}
<#
---------------------------------------------------------------------------------
Function returns seconds since 01.01.1970 00:00:00 till the DateTime in an input 
string presenting PShell DateTime string:  ddMMyyy_hhmmss  
---------------------------------------------------------------------------------
#>
function DateStringToSeconds( [string]$utc ) 
{
   $epoch = ( Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 )### CONST
   $tmp   = UTCStringToStandard( $utc )
   $ddate = [datetime]::ParseExact( $tmp, "ddMMyyyy_HHmmss", [System.Globalization.CultureInfo]::CurrentCulture )
   $ddate | % { $Seconds = [math]::truncate( $_.ToUniversalTime().Subtract( $epoch ).TotalSeconds ) }
   return $Seconds
}
function FileTimeToString( [int64]$qw )
{
   return ( [datetime]::fromFileTime( $qw ) ).ToString( "dd.MM.yyyy HH:mm:ss" )
}
<#
---------------------------------------------------------------------------------
Function converts Reg_DWORD DateTime which presents number of the seconds since 
01.01.1970 00:00:00 into QWORD FileTime number. Constant which is used in
calculations presents number of the 100 nanosecond intervals since 01.01.1601 
till 01.01.1970, 1000 - miiliseconds in 1 second, 10000 - number of the 100 ns
in 1 ms
---------------------------------------------------------------------------------
the UNIX Epoch is January 1st, 1970 at 12:00 AM (midnight) UTC
#>
function RegDWDateToFileTime( [int32]$dw )
{  
   return [int64]( ( 1000 * $dw + 11644473600000 ) * 10000 )
}
<#
-------------------------------------------------------------------------
It is impossible to add new property to the object created with
"get-childItem -Name $path" by assignment. Thus Add-Member method
is used here to create some additional properties to the input object 
-------------------------------------------------------------------------
#>
function AddProperties( [pscustomobject]$o )
{
   Add-Member -InputObject $o -type NoteProperty -name UpgradeDate -value 1427802253
   Add-Member -InputObject $o -type NoteProperty -name UpgradeDateString -value ""
   Add-Member -InputObject $o -type NoteProperty -name InstallDateString -value ""
   Add-Member -InputObject $o -type NoteProperty -name Minutes -value 777 
   return $o
}
function Show( $arr )
{
   $ShowProps = @(
   "CurrentBuild", "BuildLabEx", "InstallDateString", "UpgradeDateString", "Minutes", `
   "InstallDate", "InstallTime", "UpgradeDate", "ProductName", "ProductId", "EditionID", `
   "CurrentVersion", "ReleaseId", "UBR", "PSParentPath", "PSChildName" 
   ) 
   $MaxL = "InstallDateString".Length 
   Write-Host( "" ) > .\PS-BuildsHistory.log    
   for ( $i = 0; $i -lt $arr.Length; $i++ )
   {
      $arr[ $i ] | Select $ShowProps >> .\PS-BuildsHistory.log
   }
   return 0
}
############################# Builds objects collecting starts ######################
### EA: ErrorAction
$Omit = @( "BuildGUID","BuildLab","CompositionEditionID","CurrentBuildNumber",`
"CurrentMajorVersionNumber","CurrentMinorVersionNumber","CurrentType","Customizations",`
"DigitalProductId","DigitalProductId4","EditionID","InstallationType","PathName",`
"RegisteredOrganization","RegisteredOwner","SoftwareType","SystemRoot","MigrationScope" )

$Need = @( "BuildBranch","BuildLabEx","CurrentBuild","CurrentVersion","EditionID",`
"InstallDate","InstallTime","ProductId","ProductName","ReleaseId","SoftwareType","UBR" )

$temp  = ( get-itemproperty HKLM:\SYSTEM\Setup\"Source OS*" -Name $Need -EA SilentlyContinue ) 
$temp += ( get-itemproperty HKLM:\SOFTWARE\Microsoft\"Windows NT\CurrentVersion" -Name $Need )                     

# 
## In this loop we update some of the properties for all of the Build objects
## We can not yet count installations durations values as an array is
## not sorted by the date of installation
#
$MaxL  = "Windows 10 Pro Technical Preview".Length
$DateL = "13.10.2016 21:23:31".Length

################################################################################################
## DEVELOPER TRICK for the one of his Notebooks                                               ##
## Verify if we are running at developer computer:                                            ##
## in this case will need to correct time zone for the string dates                           ##
## less then 14.01.2016 as all previous dates were saved in the                               ##
## time zone which differs current one.                                                       ##
##                                                                                            ##
$p = "HARDWARE\DESCRIPTION\System\MultifunctionAdapter\0\DiskController\0\DiskPeripheral\0\"; ##
$MyId = (get-itemproperty -path HKLM:\$p ).Identifier# get 1-st disk unique ID                ##         
$R500 = ( $MyId -eq "9f9c4a4c-80d01f2a-A" ); ## is true until developer SSD is alive...       ## 
################################################################################################

for ( $i = 0; $i -lt $temp.Length; $i++ )
{
   $obj = AddProperties( [pscustomobject]$temp[ $i ] )

   if ( $obj.InstallTime -eq $null )## BuildNumber less them 9841
   {
      Add-Member -InputObject $obj -type NoteProperty -name InstallTime -value 1427802253   
      $obj.InstallTime = RegDWDateToFileTime( $obj.InstallDate ) 
   } 
 
   $obj.InstallDateString = FileTimeToString( $obj.InstallTime )
   
   if ( $obj.PSChildName.Substring( 0, 3 ) -eq "Sou" )## $obj.CurrentBuild should be left aligned!!!
   {
      $p = $obj.PSChildName.Replace( "Source OS (Updated on ", "" ).Replace( ")", "" )
      $obj.UpgradeDate = DateStringToSeconds( $p )
      ##########################################################
      if ( $R500 -AND [int32]($obj.CurrentBuild ) -lt 10586 )  # All upgrades till installation
      {                                                        # of the BuildNumber 10586 were 
         $obj.UpgradeDate += 3600                              # run at Time Zone one hour less 
      }                                                        # then a current one       
      ##########################################################
      $obj.UpgradeDateString = FileTimeToString( RegDWDateToFileTime( $obj.UpgradeDate ) ) 
   }
   
   ### Align some fields
   if ( $obj.CurrentBuild.Length -eq 4 )
   { 
      $obj.CurrentBuild = " " + $obj.CurrentBuild 
   }
     
   if ( $obj.UBR -ne $null )
   { 
      $obj.UBR = $obj.UBR.ToString()
      if ( $obj.UBR.Length -lt 5 )
      {
         $obj.UBR = " " * ( 5 - $obj.UBR.Length ) + $obj.UBR 
      }
   }
   if ( $obj.ProductName.Length -lt $MaxL )
   {
      $obj.ProductName += " " * ( $MaxL - $obj.ProductName.Length )
   }   
   $temp[ $i ] = $obj # update Array element   
}

### SORT an array of the Build objects
$temp = ( $temp | Sort-Object -Property InstallDate )

$temp[ 0 ].Minutes = ""
$temp[ $temp.Length - 1 ].UpgradeDateString = " " * $DateL

for ( $i = 1; $i -lt $temp.Length; $i++ )
{
   $temp[ $i ].Minutes = [math]::truncate( ( 59 + $temp[ $i ].InstallDate - $temp[ $i - 1 ].UpgradeDate ) / 60 )
}

$Debug = $False ###
if ( $Debug )
{
   ("") > .\PS-Dates.log 
   $c1 = '[PSObject]@{"IDate"='
   $c2 = '"IDateS"="'
   $c3 = '"UDate"='
   $c4 = '"UDateS"="'
   
   for ( $i = 0; $i -lt $temp.Length; $i++ )
   {
          $Out = $c1 + $temp[ $i ].InstallDate + ';' + $c2 + $temp[ $i ].InstallDateString + '";';
      if ( $i -lt $temp.Length - 1 )
      {      
         $Out += $c3 + $temp[ $i ].UpgradeDate + ';' + $c4 + $temp[ $i ].UpgradeDateString + '";},'
      }
      else
      {
         $Out += "}"
      }
      ($out) >> .\PS-Dates.log      
   } 
}

$Title = "History contains " + ( $temp.Length - 1 ) + " upgrades since " + `
$temp[ 0 ].InstallDateString + " till " + $temp[ $temp.Length - 1 ].InstallDateString
$x = Show( $temp ) ### write full log

$temp | Select-Object -Property CurrentBuild, InstallDateString, UpgradeDateString, 
ProductName, BuildBranch, ReleaseId, UBR, Minutes | Out-GridView -Title $Title -Wait
