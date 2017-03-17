<#
--------------------------------------------------------------------- 
 Project: LastSetup                                               
 Version: 1.0                                                         
 FileId:  e:\Delme\BuildsHistory\BuildInfo.ps1                        
 When:    17-Mar-17                           
 ©        Oleg Kulikov, sysprg@live.ru                                
--------------------------------------------------------------------- 
 This is simplified version of the LogoLaz.ps1 which reads data
 from the c:\Windows\Panther\setupact.log and thus GetBuildData
 can be called without parameter.
---------------------------------------------------------------------
#>

Set-StrictMode -Version Latest

$style = 
[System.Collections.Hashtable]@{
   ColorScheme       = "Black";
   LabelWidth        = "105";
   MarginLeft        = "5";
   TextWidth         = "320";
   RowHeight         = "30";
   ButtonFontSize    = "20";
   FormFontSize      = "14";
   LongLinesFontSize = "10";
};

Set-Variable -Name DefaultStyle -Scope Script -Value $style;

############################ GUI Presentaion code starts ################################

Add-Type -AssemblyName "PresentationFramework"                                                                                                                                              

<#
---------------------------------------------------------------------------------
Function receives:
[string[]]$Names  - an array of the properties names  to be placed into the Form
[string[]]$Values - an array of the properties values to be placed into the Form
[string[]]$Titles - an array of the labels of the header and trailer rows 
[PSObject]$Style  - Style object which defines Geometry and colors of the form
Function returns [xml] $xaml - Form presentation XML
---------------------------------------------------------------------------------
#>
function ShowForm
{
   param(
      [string[]]$Names,
      [string[]]$Values,
      [string[]]$Titles,
      [PSObject]$Style = $DefaultStyle
   );

   $Colors = @{
      Black = @{ Dark='#FF2F4F4F'; Light='#FFD3D3D3'; Half='#FF778899'; Body='#FFDCDCDC' };
      Brown = @{ Dark="#FF8B4516"; Light="#FFFFE4C4"; Half="#FFCD853F"; Body="#FFF5DEB3" }; 
      Green = @{ Dark="#FF006400"; Light="#FF90EF90"; Half="#FF00B300"; Body="#FF98FB98" };
      Blue  = @{ Dark="#FF0000CD"; Light="#FFADD8D6"; Half="#FF6495ED"; Body="#FF8FCEFA" }; 
   };

   $Scheme = $Style.ColorScheme;
   $TopTitle = $Titles[0];
   $BotTitle = $Titles[1];

   $dark  = $Colors.$Scheme.Dark; 
   $half  = $Colors.$Scheme.Half;  
   $light = $Colors.$Scheme.Light; 
   $body  = $Colors.$Scheme.Body; 

   $LabelWidth   = $Style.LabelWidth;     # "105"
   $MarginLeft   = $Style.MarginLeft;     # "5"
   $TextWidth    = $Style.TextWidth;      # "320"
   $TextLeft     = ([int]$MarginLeft + [int]$LabelWidth + [int]$MarginLeft).ToString();
   $RowHeight    = $Style.RowHeight;      # "30"
   $ButtonHeight = $ButtonWidth = $RowHeight; # 30
   $FullHight    = [int]$MarginLeft + [int]$RowHeight;# 5 + 30
   $btnFSZ       = $Style.ButtonFontSize; # "20";
   $winFSZ       = $Style.FormFontSize;   # "14";
  
   $LabelPattern = @"
<Label Margin="$MarginLeft,%w,0,0" Content="%n" Name="lbl%n" 
Background="$half" Foreground="#FFFFFFFF"
VerticalAlignment="Top" HorizontalAlignment="Left" 
Width="$LabelWidth" Height="$RowHeight"
/>
"@;
   $TextPattern  = @"
<TextBox Margin="$TextLeft,%w,0,0" Name="txt%n" Padding="2,4,0,0" 
Background="$light" Text="%text"
VerticalAlignment="Top" HorizontalAlignment="Left" 
Width="$TextWidth" Height="$RowHeight"
/>
"@;
 
   $OffsetTop = $FullHight;#35;
   $labels    = "";
   $text      = "";

   for ( $i = 0; $i -lt $Names.Length; $i++ )
   {
      $n = $Names[$i];
      $w = $OffsetTop.ToString().PadLeft( 3, "0" );
      $labels += $LabelPattern.Replace( "%n", "$n" ).Replace( "%w", "$w" );
      $tp = $TextPattern.Replace( "%n", "$n" ).Replace( "%w", "$w" ).Replace( "%text",$Values[$i] );
      $text   += $tp;
      $OffsetTop += $FullHight;
   }

   $FormWidth  = (3*([int]$MarginLeft)+[int]$LabelWidth+[int]$TextWidth).ToString();
   $FormHeight = ($OffsetTop + $FullHight - $MarginLeft).ToString();
   $ButtonOffsetLeft = ([int]$FormWidth - [int]$ButtonWidth).ToString();
   $stretch    = "UltraCondensed";

   $xaml = ([xml]@"                                                                           
<?xml version="1.0" encoding="UTF-8"?>                                                                                                      
<!-- XAML Code - Imported from Visual Studio Express WPF Application -->                                                                    
<Window                                                                                                                                     
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"                                                                       
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"                                                                                  
    ResizeMode="NoResize" WindowStyle="None" BorderThickness="1"                                                                            
    Background="$body" WindowStartupLocation="CenterScreen"                                                                             
    SizeToContent="Width" Height="$FormHeight" Title="$TopTitle"                                                                                   
    Topmost="False" FontSize="$winFSZ" FontStretch="$stretch">                                                                                                          
                                                                                                                                            
    <Grid Margin="0,0,0,0">                                                                                                                 
        <Label                                                                                                                              
            Width="$FormWidth" Height="$RowHeight" Margin="0,0,0,0"                                                                                      
            VerticalAlignment="Top" HorizontalAlignment="Center"                                                                            
            HorizontalContentAlignment="Center" Content="$TopTitle"
            Background="$dark" Foreground="#FFFFFFFF"                                                          
            FontWeight="Bold" Name="Header"                                                                                            
        />                                                                                                                                 
        <Button                                                                                                                             
            Width="$RowHeight" Height="$RowHeight" Margin="$ButtonOffsetLeft,0,0,0" BorderThickness="0"                                                                   
            VerticalAlignment="Top" HorizontalAlignment="Left"                                                                              
            Foreground="White" Background="#FFC35A50"                                                                                       
            FontWeight="Bold" FontSize="$btnFSZ" Content="X"                                                                                     
            Cursor="Hand" ToolTip="Close Window" Name="btnExit"                                                                             
        /> 
$labels
$text 
        <Label                                                                                                                              
            Width="$FormWidth" Height="$RowHeight" Margin="0,$OffsetTop,0,0"                                                                                      
            VerticalAlignment="Top" HorizontalAlignment="Center"                                                                            
            HorizontalContentAlignment="Center" Content="$BotTitle"
            Background="$dark" Foreground="#FFFFFFFF"                                                          
            Name="Footer"                                                                                          
        />                  
    </Grid>                                                                                                                                 
</Window>                                                                                                                                   
"@); # FontWeight="Bold" 

#####################  Preparing Form to Show ####################################

    $syncHash = [Hashtable]::Synchronized(@{});                                                                                              
    $rsHash   = [Hashtable]::Synchronized(@{});                                                                                              
    $rsHash.Runspace = [RunspaceFactory]::CreateRunspace();                                                                                
    $rsHash.Runspace.ApartmentState, $rsHash.Runspace.ThreadOptions = 'STA','ReuseThread'                                                  
    $rsHash.Runspace.Open();                                                                                                               
    $rsHash.Runspace.SessionStateProxy.SetVariable( "syncHash",$syncHash );                                                                  
    $rsHash.Runspace.SessionStateProxy.SetVariable( "rsHash", $rsHash );                                                                      
    $syncHash.xaml = $xaml;
    $rform = $null;                                                                                                        

    $psCmd = [PowerShell]::Create().AddScript(
    {                                                                                                                     
        $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $syncHash.xaml;                                                     
        $Form   = [Windows.Markup.XamlReader]::Load($reader);
        $rform  = $Form;                                                                                                                
        $Form.Add_Closed(
           {                                                                                                                
              $rshash.PowerShell.Dispose();                                                                                                                                         
              [GC]::Collect();                                                                                                               
              [GC]::WaitForPendingFinalizers();                                                                                              
           }
        ); 
 
        $lblMain = $Form.FindName('Header');                                                
        $eventHandler_LeftButtonDown = [Windows.Input.MouseButtonEventHandler]{ $Form.DragMove() };                                       
        $lblMain.add_MouseLeftButtonDown($eventHandler_LeftButtonDown);
                                                                                                                              
        $btnExit = $Form.FindName('btnExit');                                                                                                       
        $btnExit.add_Click({ $Form.Close() });                                                                                             
                                                                                                                                                                                                       
        $Form.ShowDialog();
                                                                                                             
    });# End-of-$psCmd = [PowerShell]::Create().AddScript(                                                                                                                                    
                                                                                                                                          
    $psCmd.Runspace = $rsHash.Runspace;                                                                                                    
    [void]$psCmd.BeginInvoke();                                                                                                            
    return $rform;
}

#################### Application code starts ##########################

function GetBuildData( [string]$path = "c:\Windows\Panther\setupact.log" )                                                      
{
   function SaveOldOS( [string]$S, [PSObject]$obj )
   {
        # $key02 = "Host OS Build String";        # Host OS Build String    [ 14379.0.amd64fre.rs1_release.160627-1607 ]
        $OldOS = $S.split(" ", [StringSplitOptions]::RemoveEmptyEntries)[-2];
        $obj.OldOS = "$OldOS";

        $xnn = $OldOS.Split(".")[2];
        $obj.NewOS = $obj.NewOS.Replace( "xnn", "$xnn" );# Write-Host "SaveOldOS: NewOS="$obj.NewOS;
        return $obj;
   } 
   function SaveNewOS( [string]$S, [PSObject]$obj )
   {
        # $key01 = Host OS Build String    [ 15025.1000.amd64fre.rs_prerelease.170127-1750 ] 
        $obj.NewOS = $S.trim().Substring(5).Replace( " (", ".xnn." ).Replace( ")", "" );
        return $obj;
   } 
   function SaveNewOS_( [string]$S, [PSObject]$obj )
   {
        # $key01 = "Setup build version is:";     # Setup build version is: 10.0.14385.0 (rs1_release.160706-1700) 
        $obj.NewOS = $S.trim().Substring(5).Replace( " (", ".xnn." ).Replace( ")", "" );
        return $obj;
   } 
   function SaveEdition( [string]$S, [PSObject]$obj )
   {
        # $key03 = "Host OS Edition";             # Host OS Edition         [ Professional ]
        $obj.Edition = $S.split(" ",[StringSplitOptions]::RemoveEmptyEntries)[-2];
        return $obj;
   }
   function SaveRAM( [string]$S, [PSObject]$obj )
   {
        # $key04 = "Total memory :";              # Total memory : 8589934592
        $RAM = [int64]($S.trim());
        $obj.RAM = [string]($RAM/1024/1024/1024)+"Gb";
        return $obj; 
   } 
   function SaveChannel( [string]$S, [PSObject]$obj )
   {
        # $key05 = "InstallChannel =";            # Product InstallChannel = Retail
        $obj.Channel = $S.trim();
        $obj.Edition += ", "+$obj.Channel;
        return $obj;
   }
   function SaveManufac( [string]$S, [PSObject]$obj )
   {
        # $key06 = "Manufacturer:";               # Manufacturer: LENOVO
        $obj.Manufac = $S.trim();
        return $obj;
   }
   function SaveModel( [string]$S, [PSObject]$obj )
   {
        # $key07 = "Model : ";                    # Model : 2732W12
        $obj.Model = $S.trim();
        return $obj;
   } 
   function SaveBIOSV( [string]$S, [PSObject]$obj )
   {
        # $key08 = "BIOS version :";              # BIOS version : 7YET83WW (3.13 )
        $obj.BIOSV = $S.trim();
        return $obj;
   }
   function SaveStone( [string]$S, [PSObject]$obj )
   {
        # $key09 = "Processor name :";            # Processor name : Intel(R) Core(TM)2
        $obj.Stone = $S.trim();
        return $obj;
   }
   function SaveInstall( [string]$S, [PSObject]$obj )
   {
        # $key10 = "InstallFilePath        = [C"; # InstallFilePath        = [C:\$WINDOWS.~BT\Sources\Install.esd]
        $obj.Install = "C$S".Substring(0, -1 );
        return $obj;
   }
   function SaveWimPath( [string]$S, [PSObject]$obj )
   {
        # $key11 = "MOUPG      WIM Path";         # WIM Path        [ C:\$WINDOWS.~BT\Sources\Install.esd ]
        $obj.WimPath = $S.Split(" ",[StringSplitOptions]::RemoveEmptyEntries)[-2];
        return $obj;
   }
   function SaveESDOpen( [string]$S, [PSObject]$obj )
   {
        # $key12 = "Opening ESD package file:";   # Opening ESD package file: [C:\$WINDOWS.~BT\Sources\Install.esd]
        $obj.ESDOpen = $S.trim();
        return $obj;
   }
   function SaveESDName( [string]$S, [PSObject]$obj )
   {
        #$key13 = # /WUCachedFileName "14385.0.160706-1700.rs1_release_CLIENTPRO_RET_x64fre_en-us.esd]" 
        $w = $S.Replace('"',"").Split(" ",[StringSplitOptions]::RemoveEmptyEntries);
        $obj.ESDName = ($w[0]).Split("`]")[0];
        $obj.ESDGUID = $w[-1];
        return $obj;
   }
   function SaveDataSrc( [string]$S, [PSObject]$obj )
   {
        # $key14 "External Download Available", $key15 "Install.wim" 
        $obj.DataSrc = "Network";
        if ( $S.Contains("Install.wim")){$obj.DataSrc = "ISO";}
        return $obj;
   }
   function SaveStartWU( [string]$S, [PSObject]$obj )
   {
        # $key16 "= [/PreDownload /Package /Quiet  /progressCLSID f1851d8e-504f-48a9-acf7-a8c7ff709abe 
        # /ReportId 641F4EE3-5BDA-49F6-9521-5E2E974D312A.1 /FlightData "NG:DAFA" 
        # "/CancelId" "15b6cee7-b1bb-4a1b-a5af-167fb1401ad7" 
        $obj.StartWU = $S.Substring(19);
        if ( $S.Contains("DownloadSizeInMB")){$obj.DataSrc = "Network.UUP";}
        else                                 {$obj.DataSrc = "Network.ESD";}
        return $obj;
   }
   function cs( [PSObject]$R, [string]$S, [string]$sk, [string]$kn )
   {
      if( $S.Contains( $sk )){$t = ($S -Split $sk)[1]}
      else
      {
         Write-Host "CS failed at: Key=$sk`nsrc=$S";
      }
      switch ($kn) 
      {
         "RAM"     { $R = SaveRAM     $t $R; $x = $R.$kn; return "$kn,$x"; } # $key04 "Total memory :"                             
         "Model"   { $R = SaveModel   $t $R; $x = $R.$kn; return "$kn,$x"; } # $key07 "Model : "                              
         "NewOS"   { $R = SaveNewOs   $t $R; $x = $R.$kn; return "$kn,$x"; } # $key01 "Setup build version is:"                             
         "OldOS"   { $R = SaveOldOs   $t $R; $x = $R.$kn; return "$kn,$x"; } # $key02 "Host OS Build String"                            
         "BIOSV"   { $R = SaveBIOSV   $t $R; $x = $R.$kn; return "$kn,$x"; } # $key08 "BIOS version :"                             
         "Stone"   { $R = SaveStone   $t $R; $x = $R.$kn; return "$kn,$x"; } # $key09 "Processor name :"                            
         "Edition" { $R = SaveEdition $t $R; $x = $R.$kn; return "$kn,$x"; } # $key03 "Host OS Edition"                            
         "Channel" { $R = SaveChannel $t $R; $x = $R.$kn; return "$kn,$x"; } # $key05 "InstallChannel ="                            
         "Manufac" { $R = SaveManufac $t $R; $x = $R.$kn; return "$kn,$x"; } # $key06 "Manufacturer:"                            
         "Install" { $R = SaveInstall $t $R; $x = $R.$kn; return "$kn,$x"; } # $key10 "InstallFilePath        = [C"                           
         "WimPath" { $R = SaveWimPath $t $R; $x = $R.$kn; return "$kn,$x"; } # $key11 = "MOUPG      WIM Path"                           
         "ESDOpen" { $R = SaveESDOpen $t $R; $x = $R.$kn; return "$kn,$x"; } # $key12 "Opening ESD package file:"
         "ESDName" { $R = SaveESDName $t $R; $x = $R.$kn; return "$kn,$x"; } # $key13 "/WUCachedFileName" 
         "DataSrc" { $R = SaveDataSrc $S $R; $x = $R.$kn; return "$kn,$x"; } # $key14 "External Download Available" 
                                                                             # $key15 "Install.wim" 
         "StartWU" { $R = SaveStartWU $S $R; $x = $R.$kn; return "$kn,$x"; } # $key16 "CmdLine = [/PreDownload /Package"                                                                                                       
         default {}
      }# end-of-switch
      return "$kn,"; 
   }  

   $RunLog = [string[]]@("`n-----------`n$log");
   $StampL = "2017-02-08 22:39:27".Length;
   $head   = get-content -TotalCount 5000 -Path $log; # head lines
   $tail   = get-content -Tail          2 -Path $log; # last lines
   $SStart = $head[0].Substring( 0, $StampL );        # Setup start timestamp
   $SEnd   = $tail[1].Substring( 0, $StampL );        # Setup end   timestamp
   $DStart = $DEnd = $OldOS = $NewOS = "";
   $R      = [PSObject]@{ 
               OldOS         = "";
               NewOS         = "";
               DataSrc       = "";
               StartWU       = "";
               SetupStart    = "$SStart";
               SetupEnd      = "$SEnd";
               RAM           = "";
               Channel       = "";
               Edition       = "";
               Manufac       = "";
               Model         = "";
               BIOSV         = "";
               Stone         = "";
               Install       = "";
               WimPath       = "";
               ESDOpen       = "";
               DownloadStart = "";
               DownloadEnd   = "";
               ESDName       = "";
               ESDGUID       = "";
   };
   # 2017-03-03 23:20:03, Info                  SP     Executing download operation: Download Critical Dynamic Updates
   # 2017-03-03 23:20:13, Info                  SP     Executing download operation: Download OS Updates (DU) to keep installation up-to-date
   # 2017-03-03 23:20:46, Info                  SP     Executing download operation: Download Driver Dynamic Updates
   # 2017-03-03 23:21:20, Info                  SP     Executing download operation: Download Language Pack Dynamic Updates
   $key01 = "Setup build version is:";     # Setup build version is: 10.0.14385.0 (rs1_release.160706-1700)
   $key02 = "Host OS Build String";        # Host OS Build String    [ 14379.0.amd64fre.rs1_release.160627-1607 ]   
   $key03 = "Host OS Edition";             # Host OS Edition         [ Professional ]
   $key04 = "Total memory :";              # Total memory : 8589934592         
   $key05 = "InstallChannel =";            # Product InstallChannel = Retail      
   $key06 = "Manufacturer:";               # Manufacturer: LENOVO          
   $key07 = "Model : ";                    # Model : 2732W12               
   $key08 = "BIOS version :";              # BIOS version : 7YET83WW (3.13 )         
   $key09 = "Processor name :";            # Processor name : Intel(R) Core(TM)2 Duo CPU     P8600  @ 2.40GHz       
   $key10 = 'InstallFilePath        = `[C';# InstallFilePath        = [C:\$WINDOWS.~BT\Sources\Install.esd]              
   $key11 = "MOUPG      WIM Path";         # WIM Path        [ C:\$WINDOWS.~BT\Sources\Install.esd ]   
   $key12 = "Opening ESD package file:";   # Opening ESD package file: [C:\$WINDOWS.~BT\Sources\Install.esd]
   $key13 = "/WUCachedFileName";           # /WUCachedFileName "14385.0.160706-1700.rs1_release_CLIENTPRO_RET_x64fre_en-us.esd" 
                                           # /SuccessId 7ed0e81d-29bb-4f7f-9fa8-394f1bb7e39e]
   $key14 = "External Download Available"; # MOUPG  SetupHost: External Download Available
   $key15 = "Install.wim";                 # MOUPG  SetupHost::Initialize: InstallFilePath = [D:\Sources\Install.wim] 
   $key16 = "/PreDownload";                # CmdLine                = [/PreDownload 

   for ( $i = 1; $i -lt $head.Length; $i++ )
   {
       $S = $head[$i];
       $p = "";
       # Product InstallChannel = 
       if     ( !$R.DataSrc -and $S.Contains( "$key14" ) ){ $p=cs $R $S $key14 "DataSrc"; }
       elseif ( !$R.DataSrc -and $S.Contains( "$key15" ) ){ $p=cs $R $S $key15 "DataSrc"; }
       elseif ( !$R.StartWU -and $S.Contains( "$key16" ) ){ $p=cs $R $S $key16 "StartWU"; }
       elseif ( !$R.NewOS   -and $S.Contains( "$key01" ) ){ $p=cs $R $S $key01 "NewOS";   }                                                               
       elseif ( !$R.OldOS   -and $S.Contains( "$key02" ) ){ $p=cs $R $S $key02 "OldOS";   }                                                               
       elseif ( !$R.Edition -and $S.Contains( "$key03" ) ){ $p=cs $R $S $key03 "Edition"; }                                                               
       elseif ( !$R.RAM     -and $S.Contains( "$key04" ) ){ $p=cs $R $S $key04 "RAM";     }                                                               
       elseif ( !$R.Channel -and $S.Contains( "$key05" ) ){ $p=cs $R $S $key05 "Channel"; }                                                               
       elseif ( !$R.Manufac -and $S.Contains( "$key06" ) ){ $p=cs $R $S $key06 "Manufac"; }                                                               
       elseif ( !$R.Model   -and $S.Contains( "$key07" ) ){ $p=cs $R $S $key07 "Model";   }                                                               
       elseif ( !$R.BIOSV   -and $S.Contains( "$key08" ) ){ $p=cs $R $S $key08 "BIOSV";   }                                                               
       elseif ( !$R.Stone   -and $S.Contains( "$key09" ) ){ $p=cs $R $S $key09 "Stone";   }                                                               
       elseif ( !$R.Install -and $S.Contains( "$key10" ) ){ $p=cs $R $S $key10 "Install"; }                                                               
       elseif ( !$R.WimPath -and $S.Contains( "$key11" ) ){ $p=cs $R $S $key11 "WimPath"; }                                                               
       elseif ( !$R.ESDOpen -and $S.Contains( "$key12" ) ){ $p=cs $R $S $key12 "ESDOpen"; }
       elseif ( !$R.ESDName -and $S.Contains( "$key13" ) ){ $p=cs $R $S $key13 "ESDName"; }                                                             
       elseif ( $s.Contains( "Power request granted!"  ) )  
       {
           $nline = $S = $head[ 1 + $i ];
           if ( $nline.IndexOf( "SetupPhasePreDownload" ) -gt -1 )
           {
               $R.DownloadStart = $nline.Substring( 0, $StampL );
               $x = $R.DownloadStart;
               $p = "DownloadStart,$x";
           }
           elseif ( $R.DataSrc.StartsWith("Network") -and $nline.IndexOf( "SetupPhaseUnpack" ) -gt -1 )
           {
               $R.DownloadEnd = $nline.Substring( 0, $StampL );
               $p = $R.DownloadEnd;
               $p = "DownloadEnd,$p";
           }
       }
       if ( $p )############## Debug code ################
       { #2016-07-10 21:24:13
         $j = (1+$i).ToString().PadLeft(4,"0"); 
         $t = $S.Substring(11,8);
         $w = $p.Split(",")#.PadRight( 13, " " );#DownloadStart
         $OffsetTop = ($w[0]).PadRight( 13, " " );
         $w1 = $w[1];
         $RunLog+="$j,$t,$OffsetTop,$w1"; 
       }
   }# End-of-I-loop

   $ret = [PSObject]@{             #   $R      = [PSObject]@{                       
               OldOSMedia   = "";  #               OldOS         = "";              
               NewOSMedia   = "";  #               NewOS         = "";              
               DataSource   = "";  #               DataSrc       = "";              
               StartWU      = "";  #               StartWU       = "";              
               SetupStart   = "";  #               SetupStart    = "$SStart";       
               SetupEnd     = "";  #               SetupEnd      = "$SEnd";         
               RAM          = "";  #               RAM           = "";              
               Channel      = "";  #               Channel       = "";              
               Edition      = "";  #               Edition       = "";              
               Manufacturer = "";  #               Manufac       = "";              
               Model        = "";  #               Model         = "";              
               BIOSVer      = "";  #               BIOSV         = "";              
               Processor    = "";  #               Stone         = "";              
               Install      = "";  #               Install       = "";              
               WimPath      = "";  #               WimPath       = "";              
               ESDOpen      = "";  #               ESDOpen       = "";              
               DownloadStart = ""; #               DownloadStart = "";              
               DownloadEnd   = ""; #               DownloadEnd   = "";              
#               ESDName       = ""; #               ESDName       = "";              
#               ESDGUID       = ""; #               ESDGUID       = "";              
   };                              #   };                                           


   $ret.OldOSMedia   = $R.OldOS;
   $ret.NewOSMedia   = $R.NewOS;
   $ret.DataSrc      = $R.DataSrc;
   $ret.SetupStart   = $R.SetupStart;
   if ( $R.DataSrc.StartsWith( "Network" ) )
   {
      $ret.DownloadStart = $R.DownloadStart; 
      $ret.DownloadEnd = $R.DownloadEnd;
      if ( $R.DataSrc.EndsWith( "ESD" ) ) {$ret.ESDName=$R.ESDName;}  
   }
   $ret.SetupEnd     = $R.SetupEnd;
   $ret.RAM          = $R.RAM;
   $ret.Channel      = $R.Channel;
   $ret.Edition      = $R.Edition;
   $ret.Processor    = $R.Stone;
   $ret.Install      = $R.Install;
   $ret.WimPath      = $R.WimPath;
   $ret.ESDOpen      = $R.ESDOpen;
   $ret.Manufacturer = $R.Manufac;
   $ret.Model        = $R.Model;
   $ret.BIOSVer      = $R.BIOSV;
   $ret.RAM          = $R.RAM;
   $ret.Platform     = ( $R.Manufac + " " + $R.Model + " " + $R.BIOSV ).trim();
   $ret.ESDName      = $R.ESDName;
   $R = $null;

   $nams = "OldOSMedia,NewOSMedia,Edition,DataSrc,SetupStart,";


   if ( $ret.DataSrc.StartsWith("Network" ) )
   {
      $nams += "DownloadStart,DownloadEnd,";
   }
   $nams += "SetupEnd,Platform,RAM,Processor";
   if ( $ret.DataSrc.EndsWith( "ESD" ) ) {$nams+=",ESDName";}
   return @{ Names=$nams.Split(","); Object=$ret; Stamps=$RunLog };                                                                                                                                         
}#-end-of-GetBuildData
$log   = "c:\Windows\Panther\setupact.log";
$rlog  = [string[]]@();
$tmp   = GetBuildData $log;
$rlog += $tmp.Stamps; 
$l     = $rlog.Length; 
$Obj   = [PSObject]$tmp.Object;
$Names = [String[]]$tmp.Names;
$IsUUP = ($Obj.DataSrc).Contains("UUP");
$IsISO = ($Obj.DataSrc).Contains("ISO");
$Vals  = [string[]]@();
for ( $j = 0; $j -lt $Names.Length; $j++ )
{
   $p = $Names[$j];
   $Vals += $Obj.$p;
}
$Heads = [string[]]@( "Windows Setup report","$log" );
$style = 
[System.Collections.Hashtable]@{
   ColorScheme       = "Black";
   LabelWidth        = "95";
   MarginLeft        = "3";
   TextWidth         = "380";
   RowHeight         = "26";
   ButtonFontSize    = "18";
   FormFontSize      = "12";
   LongLinesFontSize = "10";
};

if     ( $IsUUP ){$style.ColorScheme = "Brown";} 
elseif ( $IsISO ){$style.ColorScheme = "Green";} 
else             {$style.ColorScheme = "Blue";} 
$xaml = ShowForm $Names $Vals $Heads $style; 

# https://www.microsoft.com/en-au/download/details.aspx?id=54794
<#--End-of-File-LastSetup----------------------------------------#>


