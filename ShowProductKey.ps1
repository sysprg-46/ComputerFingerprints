<#
--------------------------------------------------------------------- 
 Project: GetProductKey                                               
 Version: 1.0                                                         
 FileId:  e:\Delme\BuildsHistory\GetProductKey                        
 When:    22 Oct 2016,  Saturday,  14:27:12                           
 Who:     Oleg Kulikov, sysprg@live.ru                                
--------------------------------------------------------------------- 
 This is an upgrade of an old code created for Windows 7, Windows 8:
 Code below follows get_windows_8_key.vbs presented at
 http://winitpro.ru/index.php/2012/10/12/kak-uznat-klyuch-windows-8/
 with insignificant changes
 OlegKulikov46@mail.ru, 17-12-2012
--------------------------------------------------------------------- 
This is an attempt to produce a code which will correctly translate
Registry data to Windows 10 ProductKey. Correctness can be verified
by partial key presented by slmgr.vbs.
---------------------------------------------------------------------
23.10.2016 As SLMGR.VBS was not changed since 8.1 then algorithm 
which is used by function GetPK and DigitalProductId4 data usage
has no any sense. 
--------------------------------------------------------------------- 
13.11.2016, an algorithm is revised: shift operations are
implemented.
---------------------------------------------------------------------
14.11.2016, GUI interface implemented on base of:
http://poshcode.org/5429, Get-WindowsProductGUI by skourlatov
---------------------------------------------------------------------
#>

Add-Type -AssemblyName 'PresentationFramework'                                                 
                                                                                               
Function GetProductKey                                                      
{ 
   function ByteArrayToString( [byte[]]$arr )
   {
      $s = [System.BitConverter]::ToString( $arr );
      if ( "00" -eq $s.Substring( 0, 2 ) ){ return ""; }
      if ( "00" -eq $s.Substring( 3, 2 ) )
      {
         $a = $s.Substring( 0, $s.indexOf( "-00-00-" ) ).Replace( "-00-", "," ).split(",");
         $v = "";
         for ( $i = 0; $i -lt $a.Length; $i++ )
         {
            $v += [char]( [byte]( "0x"+$a[$i] ) );
         }
         return $v;
      }
      $a = $s.Substring( 0, $s.indexOf( "-00" ) ).split( "-" );
      $v = "";
      for ( $i = 0; $i -lt $a.Length; $i++ )
      {
         $v += [char]( [byte]( "0x"+$a[$i] ) );
      }
      return $v;
   }
                                                                                             
   function DecryptWindowsKey( [byte[]]$ByteArray )
   {
      $TrTable = "BCDFGHJKMPQRTVWXY2346789"
      $TrTabL  = $TrTable.Length;
      $RawData = $ByteArray;
      $Last    = $RawData.Length - 1;
      
      [bool]$isWin8 = ( ( $RawData[ $Last ] -shr 3 ) -band 1 );                                                           
      $RawData[ $Last ] = $RawData[ $Last ] -band 0xF7; 
       
      $PK = "";                                                                                    
      for ( $i = $TrTabL; $i -ge 0; $i-- )                                                                   
      {                                                                                                 
         $TrTabIndex = 0;                                                                                        
         for ( $j = $Last; $j -ge 0; $j-- )                                                                
         {
            $TrTabIndex = $TrTabIndex -shl 8 -bxor $RawData[ $j ];                                                      
            $RawData[ $j ] = [Math]::Truncate( $TrTabIndex / $TrTabL );            
            $TrTabIndex = $TrTabIndex % $TrTabL            
         }                                                                                              
         $PK = $TrTable[ $TrTabIndex ] + $PK;                                                                    
      }                                                                                                                                                                                        
                                                                                                        
      if ( $isWin8 )                                                                              
      {
         $i = $TrTable.indexOf( $PK.Substring( 0, 1 ) );                                                                                                                                       
         $PK = $PK.Substring( 1 ).Insert( $i,'N' );                                                                                  
      }                                                                                                 
                                                                                                        
      return "" +`                                                                                      
             $PK.Substring(  0, 5 ) + "-" +`                                                            
             $PK.Substring(  5, 5 ) + "-" +`                                                            
             $PK.Substring( 10, 5 ) + "-" +`                                                            
             $PK.Substring( 15, 5 ) + "-" +`                                                            
             $PK.Substring( 20, 5 );                              
   }#End-of-DecryptWindowsKey                                                                                                                                       
   
   $Comp    = Get-ItemProperty -Path "HKLM:\HARDWARE\DESCRIPTION\System\BIOS"; 
   $stone   = Get-ItemProperty -Path "HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\0";   
   $Current = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion"; 
   $IE      = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Registration";

   if ( $Current.InstallTime -eq $null )## BuildNumber is less then 9841
   {
      $it = [int64]( ( 1000 * $Current.InstallDate + 11644473600000 ) * 10000 );
      Add-Member -InputObject $Current -type NoteProperty -name InstallTime -value $it; 
   }

   [string]$CompID = "";
   if ( $Comp.BaseBoardManufacturer  -ne $null ) { $CompID += $Comp.BaseBoardManufacturer + " "; }
   elseif ( $Comp.SystemManufacturer -ne $null ) { $CompID += $Comp.SystemManufacturer    + " "; }
   if ( $Comp.SystemVersion          -ne $null ) { $CompID += $Comp.SystemVersion         + " "; }
   elseif ( $Comp.SystemFamily       -ne $null ) { $CompID += $Comp.SystemFamily          + " "; }  
   if ( $Comp.SystemProductName      -ne $null ) { $CompID += $Comp.SystemProductName; }

   $arch     = switch ([ Environment]::Is64BitOperatingSystem ) { $true {"x64"}; $false {"x86"} };
   $Edition  = $Current.EditionID+", $arch";    
   $PK       = DecryptWindowsKey( $Current.DigitalProductId[52..66] ); 
   $Channel  = ByteArrayToString( $Current.DigitalProductId4[0x3f8..0x478] );
   $IEPK     = DecryptWindowsKey( $IE.DigitalProductId[52..66] );     

   return [PSCustomObject]@{ 
       InstallMedia  = $Current.BuildLabEx;
       ProductName   = $Current.ProductName;          
       Version       = [Environment]::OSVersion.VersionString;                                                                             
       Edition       = $Edition;                              
       ProductID     = $Current.ProductId; 
       Installed     = [datetime]::fromFileTime( $Current.InstallTime ).ToString( "dddd, dd.MM.yyyy HH:mm:ss" );      
       ProductKey    = $PK; 
       Channel       = $Channel; 
       IEProductId   = $IE.ProductId;
       IEProductKey  = $IEPK;
       Platform      = $CompID;
       Processor     = $stone.ProcessorNameString;#.Replace( /`s/g, "_" ).replace( /__*/g, " " );      
   }                                                                                                                                       
}#-end-of-GetProductKey

$xaml = ([xml]@"                                                                           
<?xml version="1.0" encoding="UTF-8"?>                                                                                                      
<!-- XAML Code - Imported from Visual Studio Express WPF Application -->                                                                    
<Window                                                                                                                                     
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"                                                                       
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"                                                                                  
    ResizeMode="NoResize" WindowStyle="None" BorderThickness="0"                                                                            
    Background="#FFF0F0F0" WindowStartupLocation="CenterScreen"                                                                             
    SizeToContent="Width" Height="470" Title="Current Operating System Details"                                                                                   
    Topmost="False" FontSize="14">                                                                                                          
                                                                                                                                            
    <Grid Margin="0,0,0,0">                                                                                                                 
        <Label                                                                                                                              
            Width="440" Height="30" Margin="0,0,0,0"                                                                                        
            VerticalAlignment="Top" HorizontalAlignment="Center"                                                                            
            HorizontalContentAlignment="Center" Content="Current Operating System Details"                                                          
            FontWeight="Bold" Name="mainLabel"                                                                                              
        />                                                                                                                                  
        <Button                                                                                                                             
            Width="30" Height="30" Margin="410,0,0,0" BorderThickness="0"                                                                   
            VerticalAlignment="Top" HorizontalAlignment="Left"                                                                              
            Foreground="White" Background="#FFC35A50"                                                                                       
            FontWeight="Bold" FontSize="20" Content="X"                                                                                     
            Cursor="Hand" ToolTip="Close Window" Name="btnExit"                                                                             
        /> 
        <Label Margin="0,035,0,0" Content="InstallMedia"  Name="lblInstallMedia"/>
        <Label Margin="0,070,0,0" Content="ProductName"   Name="lblProductName"/>        
        <Label Margin="0,105,0,0" Content="Version"       Name="lblVersion"/>                                                                      
        <Label Margin="0,140,0,0" Content="Edition"       Name="lblEdition"/>                                                                         
        <Label Margin="0,175,0,0" Content="Product ID"    Name="lblProductID"/> 
        <Label Margin="0,210,0,0" Content="Installed"     Name="lblInstalled"/>           
        <Label Margin="0,245,0,0" Content="Product Key"   Name="lblProductKey"/>                                                              
        <Label Margin="0,280,0,0" Content="Channel"       Name="lblChannel"/>   
        <Label Margin="0,315,0,0" Content="IE ProductID"  Name="lblIEProductID"/>                                                              
        <Label Margin="0,350,0,0" Content="IE ProductKey" Name="lblIEProductKey"/>  
        <Label Margin="0,385,0,0" Content="Computer"      Name="lblPlatform"/>
        <Label Margin="0,420,0,0" Content="Processor"     Name="lblProcessor"/>          
               
        <TextBox Margin="115,035,0,0"  Name="txtInstallMedia"/>                                                                               
        <TextBox Margin="115,070,0,0"  Name="txtProductName"/>                                                                                
        <TextBox Margin="115,105,0,0"  Name="txtVersion"/>                                                                                   
        <TextBox Margin="115,140,0,0"  Name="txtEdition"/>                                                                                    
        <TextBox Margin="115,175,0,0"  Name="txtProductID"/> 
        <TextBox Margin="115,210,0,0"  Name="txtInstalled"/>         
        <TextBox Margin="115,245,0,0"  Name="txtProductKey"/>       
        <TextBox Margin="115,280,0,0"  Name="txtChannel"/> 
        <TextBox Margin="115,315,0,0"  Name="txtIEProductID"/>       
        <TextBox Margin="115,350,0,0"  Name="txtIEProductKey"/> 
        <TextBox Margin="115,385,0,0"  Name="txtPlatform"/>
        <TextBox Margin="115,420,0,0"  Name="txtProcessor"/>        
    </Grid>                                                                                                                                 
</Window>                                                                                                                                   
"@);# end-of-$xaml

#
### You can't change or delete Constants without closing PowerShell.
### https://powershell.org/forums/topic/constants-in-powershell/ 
### Set-Variable -Force -Value 
#
if ( $prod_xaml -eq $null )
{ 
   New-Variable -Option "READONLY" -Name 'prod_xaml' -Value $xaml;
}
else
{
   Set-Variable -Name 'prod_xaml' -Value $xaml -Force; 
}
                                                                                                                                           
Function ShowCurrentBuild                                                                                                              
{                                                                                                                                           
    $syncHash = [Hashtable]::Synchronized(@{});                                                                                              
    $rsHash   = [Hashtable]::Synchronized(@{});                                                                                              
    $rsHash.Runspace = [RunspaceFactory]::CreateRunspace();                                                                                
    $rsHash.Runspace.ApartmentState,$rsHash.Runspace.ThreadOptions = 'STA','ReuseThread'                                                  
    $rsHash.Runspace.Open();                                                                                                               
    $rsHash.Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash);                                                                  
    $rsHash.Runspace.SessionStateProxy.SetVariable("rsHash",$rsHash);                                                                      
    $syncHash.xaml = $prod_xaml;                                                                                                           
                                                                                                                   
    $os = GetProductKey                                                                                                              
    $syncHash.Data = New-Object string[](12)
    $syncHash.Data[0]  = $os.InstallMedia;
    $syncHash.Data[1]  = $os.ProductName;     
    $syncHash.Data[2]  = $os.Version;                                                                                                   
    $syncHash.Data[3]  = $os.Edition;                                                                                                    
    $syncHash.Data[4]  = $os.ProductId; 
    $syncHash.Data[5]  = $os.Installed;     
    $syncHash.Data[6]  = $os.ProductKey;                                                                                                
    $syncHash.Data[7]  = $os.Channel;  
    $syncHash.Data[8]  = $os.IEProductID; 
    $syncHash.Data[9]  = $os.IEProductKey; 
    $syncHash.Data[10] = $os.Platform; 
    $syncHash.Data[11] = $os.Processor;

    $psCmd = [PowerShell]::Create().AddScript(
    {                                                                                                                     
        $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $syncHash.xaml;                                                     
        $Form   = [Windows.Markup.XamlReader]::Load($reader);                                                                                                                
        $Form.Add_Closed(
           {                                                                                                                
              $rshash.PowerShell.Dispose();                                                                                                                                         
              [GC]::Collect();                                                                                                               
              [GC]::WaitForPendingFinalizers();                                                                                              
           }
        );                                                                                              
        $labels  = New-Object Collections.Generic.List[Windows.Controls.Label];                                                            
        $txtblks = New-Object Collections.Generic.List[Windows.Controls.TextBox];       
        $syncHash.xaml.SelectNodes("//*[@Name]") | foreach {                                                                              
                                                                                                                                          
            $element = $Form.FindName($_.Name);                                                                                            
            switch -Regex ($_.Name)                                                                                                       
            {                                                                                                                             
                     'lbl*' {  $labels.Add($element); }                                                                                    
                     'txt*' { $txtblks.Add($element); }                                                                                    
                  'btnExit' {  $btnExit  = $element;  }                                                                                    
                'mainLabel' {  $lblMain  = $element;  }                                                                                    
            }                                                                                                                             
        } 
                                                     
        $eventHandler_LeftButtonDown = [Windows.Input.MouseButtonEventHandler]{ $Form.DragMove() }                                        
        $lblMain.add_MouseLeftButtonDown($eventHandler_LeftButtonDown)                                                                                                      
        $color_shading = {                                                                                                                
            $lblMain.Background = '#FF45A6ED'                                                                                             
            $labels  | ForEach-Object -Process { $_.Background = '#FFA5D3F5' }                                                            
            $txtblks | ForEach-Object -Process { $_.Background = '#FFFFF0BB' }                                                            
        }                                                                                                                                 
        $monochrome_shading = {                                                                                                           
            $lblMain.Background = '#FF8D8D8D'                                                                                             
            $labels  | ForEach-Object -Process { $_.Background = '#FFC7C7C7' }                                                            
            $txtblks | ForEach-Object -Process { $_.Background = '#FFF0F0F0' }                                                            
        }                                                                                                                                 
    ##    Add events to Form Objects                                                                                                      
        $btnExit.add_Click({ $Form.Close() })                                                                                             
    ##    Make the mouse act like something is happening                                                                                  
        $btnExit.add_MouseEnter({ & $monochrome_shading })                                                                                
    ##    Switch back to regular mouse                                                                                                    
        $btnExit.add_MouseLeave({ & $color_shading })                                                                                     
    ##    Initialize form 
   
        for ($i = 0; $i -lt $labels.Count; $i++)                                                                                          
        {                                                                                                                                 
            $tb,$lb = $txtblks[$i],$labels[$i]                                                                                            
                                                                                                                                          
            $tb.IsReadOnly = $true                                                                                                        
            $tb.Text = $syncHash.Data[$i]                                                                                                 
            $tb.Width,$lb.Width = 320,110                                                                                                 
            $tb.Height = $lb.Height = 30                                                                                                  
            $tb.TextWrapping = [Windows.TextWrapping]::Wrap                                                                               
            @($tb,$lb) | foreach {                                                                                                        
                $_.VerticalAlignment = [Windows.VerticalAlignment]::Top;                                                                   
                $_.HorizontalAlignment = [Windows.HorizontalAlignment]::Left;                                                              
            }                                                                                                                             
        }                                                                                                                                 
        & $color_shading                                                                                                              
        $Form.ShowDialog()                                                                                                                
    })# End-of-$psCmd = [PowerShell]::Create().AddScript(                                                                                                                                    
                                                                                                                                          
    $psCmd.Runspace = $rsHash.Runspace                                                                                                    
    [void]$psCmd.BeginInvoke()                                                                                                            
}
                                                                                                                                         
ShowCurrentBuild

Write-Host "ShowProductKey execution finished."


<#--End-of-File-ShowProductKey----------------------------------------#>


