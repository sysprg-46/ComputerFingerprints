/*
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
only by running this code at Windows 10 activated by Retail key.
---------------------------------------------------------------------
October 23. As SLMGR.VBS was not changed since 8.1 then algorithm 
which is used by function GetPK and DigitalProductId4 data usage
has no any sense. Thus BuildNLessThen10240 should return correct
--------------------------------------------------------------------- 
*/

function myStdReg()
{
   var HKLM        = 0x80000002;    
   var objLocator  = new ActiveXObject( "WbemScripting.SWbemLocator" ); 
   var objService  = objLocator.ConnectServer( ".", "root\\default" ); 
   var objReg      = objService.Get( "StdRegProv" ); 

   var objMethod   = null;
   var objInParam  = null;
   var objOutParam = null;
   var res         = null;

   var root = HKLM;
   
   this.RegEnumSubKeys = function( root, strRegPath ) 
   { 
      var res     = null; 
      objMethod   = objReg.Methods_.Item( "EnumKey" );
      objInParam  = objMethod.InParameters.SpawnInstance_();
      objInParam.hDefKey     = root;
      objInParam.sSubKeyName = strRegPath; 
        
      objOutParam = objReg.ExecMethod_( objMethod.name, objInParam );
   
      if ( objOutParam.ReturnValue == 0 ) 
      {  
         res = objOutParam.sNames == null ? null : objOutParam.sNames.toArray(); 
      }
      else
      {
         rc = StdRegFailed( "RegEnumSubKeys", objInParam, objOutParam );
      }
      return res; 
   }// End-of-this.RegEnumSubKeys
}// End-of-MyStdReg 

function GetRootBuildPath( Shell )
{
   var HKLM     = 0x80000002; 
   var myRegObj = new myStdReg();// create service object to work with Registry
   var path     = "SYSTEM\\Setup", needKey = "Source OS (Updated on ";
   var needKeyL = needKey.length;
   var p, tmp1, tmp2, tmp3;
 
   var Minimal = 14951;
   var MinimalDir = "";
   var res = myRegObj.RegEnumSubKeys( HKLM, path );

   for ( var i = 0; i < res.length; i++ ) 
   {                                        
      p = res[ i ].indexOf( needKey );  
      if ( p > - 1 )                    
      { 
         tmp1 = path + "\\" + res[ i ];
         tmp2 = "HKLM\\" + tmp1 + "\\CurrentBuild";
         tmp3 = parseInt( Shell.RegRead( tmp2 ) );
         if ( tmp3 < Minimal )
         {
            Minimal    = tmp3;
            MinimalDir = tmp1;
         }                                   
      }//End-of-if ( p > - 1 ) 
   }//End-of-i-loop
   return ( Minimal > 9600 ? "" : MinimalDir );
}

/*
---------------------------------------------------------------------
This function is somewhat consized version of the original code, 2012
and it correctly translates Windows 7, 8, 8.1 Production key which
is proved by slmgr.vbs partial Production Key. 
To simplify translation to PShell, all variables names starts with 
"$"-character.
---------------------------------------------------------------------
*/
function BuildNLessThen10240( $RegData ) 
{ 
   var $ByteArray = $RegData;
   $IsWin8 = ( ( ( $ByteArray[ 66 ] & 0xFF ) %  6 ) & 1 );
   $ByteArray[ 66 ] = ( $ByteArray[ 66 ] & 0xF7 ) | ( ( $IsWin8 & 2 ) * 4 );  
   var $TrTable = "BCDFGHJKMPQRTVWXY2346789";  
   var $PK = "";  
   for ( var $i = 24; $i >= 0; $i-- ) 
   { 
      var $r = 0; 
      for ( var $j = 14; $j >= 0; $j-- ) 
      { 
         $r = $r * 256;
         $r = $ByteArray[ $j + 52 ] + $r;
         $ByteArray[ $j + 52 ] =  Math.floor( $r / 24 ); 
         $r = $r % 24; 
      } 
      $PK = $TrTable.charAt( $r ) + $PK;
   } 
   if ( $IsWin8 == 1 )
   {
      var $t1 = $PK.substr( 1, $r );
      $PK = $PK.substr( 1 ).replace( $t1, $t1 + 'N' );
      if ( $r == 0 ) $PK = 'N' + $PK;
   }
   return "" +
          $PK.substr(  0, 5 ) + "-" + 
          $PK.substr(  5, 5 ) + "-" + 
          $PK.substr( 10, 5 ) + "-" +
          $PK.substr( 15, 5 ) + "-" + 
          $PK.substr( 20, 5 );
}
                                                                        
var objShell  = new ActiveXObject( "WScript.Shell" );
var comspec   = objShell.ExpandEnvironmentStrings( "%comspec%" );
var fso       = new ActiveXObject( "Scripting.FileSystemObject" );
var MyName    = WScript.ScriptName;
var myAbsPath = fso.GetAbsolutePathName( "." );
var rc        = false;

var objShellWindows = new ActiveXObject( "shell.application" ).Windows();
        
if ( objShellWindows != null )
{
   var $Help = "" +
   "-----------------------------------------------------------------------------\n" +
   "Veteran Insiders keys for activation RTM 10240:                              \n" + 
   "Инсайдерские ключи для активации RTM 10240:                                  \n" + 
   "-----------------------------------------------------------------------------\n" +
   "Windows 10 Pro build 10240_________________ - VK7JG-NPHTM-C97JM-9MPGT-3V66T  \n" +
   "Windows 10 Home build 10240________________ - YTMG3-N6DKC-DKB77-7M9GH-8HVX7  \n" +
   "Windows 10 Home SingleLanguage build 10240_ - BT79Q-G7N6G-PGBYW-4YWX6-6F4BT  \n" +
   "Windows 10 Home CountrySpecific build 10240 - N2434-X9D7W-8PF6X-8DV9T-8TYMD  \n" +
   "Windows 10 Enterprise build 10240__________ - XGVPP-NMH47-7TTHJ-W3FW7-8HV2C  \n" +
   "Windows 10 Pro VL build 10240______________ - QJNXR-7D97Q-K7WH4-RYWQ8-6MT6Y  \n" + 
   "-----------------------------------------------------------------------------\n" + 
   "BBBBB-BBBBB-BBBBB-BBBBB-BBBBB - means that system was not activated with OEM \n" +
   "or Retail key or OEM/Retail key was deliberatly eliminated\n" +
   "BBBBB-BBBBB-BBBBB-BBBBB-BBBBB - означает, что система фктивирована не OEM или\n" +
   "ритейл ключем или поле ключа был преднамеренно очищено\n";
   var Shell   = new ActiveXObject( "WScript.Shell" );
   var $rk1    = "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProductName";
   var $rk2    = "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProductID";
   var $rk3    = "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\DigitalProductId";
   var $rk4    = "HKLM\\SYSTEM\\CurrentControlSet\\Control\\CommonGlobUserSettings" +
                 "\\Control Panel\\International\\LocaleName";
   var $rk5    = "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\CurrentBuild";
   var $Bytes  = Shell.RegRead( $rk3 ).toArray();
   var $BuildN  = parseInt( Shell.RegRead( $rk5 ) );   
   var $IsWin10 = ( $BuildN > 9600 );
   
   var $k1     = BuildNLessThen10240( $Bytes );
   var $Msg    = "CurrentBuildNumber:   " + $BuildN + "\n" +
                 "Windows Product Name: " + Shell.RegRead( $rk1 )  + '\n' +
                 "Windows Product ID:   " + Shell.RegRead( $rk2 )  + '\n' +
                 "Windows Key:          " + $k1;

   WScript.Echo( $Msg );
   
   $Origin = GetRootBuildPath( Shell ); 
   if ( $Origin != "" )
   {
      $rk1    = "HKLM\\" + $Origin + "\\ProductName";
      $rk2    = "HKLM\\" + $Origin + "\\ProductID";
      $rk3    = "HKLM\\" + $Origin + "\\DigitalProductId";
      $rk5    = "HKLM\\" + $Origin + "\\CurrentBuild";
      $Bytes  = Shell.RegRead( $rk3 ).toArray();
      $Locale = ( Shell.RegRead( $rk4 ) ).toLowerCase().substr( 0, 2 );
      $BuildN = parseInt( Shell.RegRead( $rk5 ) );   
   
      $k1     = BuildNLessThen10240( $Bytes );
      $Msg    = "-----------------------------------------------------------------------------\n" +
                "CurrentBuildNumber:   " + $BuildN + "\n" +
                "Windows Product Name: " + Shell.RegRead( $rk1 )  + '\n' +
                "Windows Product ID:   " + Shell.RegRead( $rk2 )  + '\n' +
                "Windows Key:          " + $k1 + "\n";
      WScript.Echo( $Msg + $Help );
   }
}

//--End-of-File-GetProductKey------------------------------------------ 


