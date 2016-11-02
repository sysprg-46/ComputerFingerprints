/*
--------------------------------------------------------------------- 
Project: BuildsHistory                                               
Version: 1.2                                                         
FileId:  f:\Delme\CleanBCD\BuildsHistory                             
When:    29 May 2015,  Friday,  07:14, Windows build 10125                             
Who:     Oleg Kulikov, sysprg@live.ru                                                                                          
--------------------------------------------------------------------- 
October 2016, running at build 14936. 
New implementations:
1. Array of the build objects is sorted by the build number AND InstallDate
   to enable natural order of the objects EVEN IN THE CASE that the same
   Build was installed more then once.
2. Have implemented some logic to change default current time zone 
   to those which was used on my notebook till December 17 2015. 
3. function for string datetime presentation was fixed. 
4. List of the Build object properties read from Registry was extended and
   some additional properties were implemented.
5. Format of the output logs was changed and depends on the optional key
   in the command. 
6. notepad is started to view full log, item 5 code temporary eliminated

Friday, October 7, 2016, finished, still build 14936 (:                                                          
---------------------------------------------------------------------
9600 properties:
BuildGUID              "ffffffff-ffff-ffff-ffff-ffffffffffff"
BuildLab               "9600.winblue_r8.150127-1500"
BuildLabEx             "9600.17668.amd64fre.winblue_r8.150127-1500"
CurrentBuild           "9600"
CurrentBuildNumber     "9600"
CurrentType            "Multiprocessor Free"
CurrentVersion         "6.3"
DigitalProductId       hex:a4,00,00,00,03,00,00,00,30,30,31,37,38,2d,31,30,35,31,..
DigitalProductId4      hex:f8,04,00,00,04,00,00,00,30,00,36,00,34,00,30,00,31,00,...
EditionID              "Professional"
InstallationType       "Client"
InstallDate            dword:5511a84e
MigrationScope         dword:00000005
PathName               "C:\\Windows"
ProductId              "00178-10511-24375-AB079"
ProductName            "Windows 8.1 Pro"
RegisteredOrganization ""
RegisteredOwner        ....................
SoftwareType           "System"
SystemRoot             "C:\\Windows"
--------------------------------------------------------------------- 
10049: CurrentMajorVersionNumber dword   0000000a
       CurrentMinorVersionNumber dword   00000000
       Customizations            String  "ModernApps"
       InstallTime               qword

10122: BuildBranch String  "fbl_impressive"
       UBR         dword   0000a309

10240: run         String  "MSASCui.exe"   is absent in any other build       

10565: ReleaseId   String  "1511"

---------------------------------------------------------------------
http://forums.mydigitallife.info/threads/58404-Index-Windows-10-Builds
---------------------------------------------------------------------
*/

//
// Set of the project independent methods used by the project code
//
String.prototype.strip = function()                                         
{                                                                           
	return this.replace( /^(\s*)/, "" ).replace( /(\s*)$/, "" );               
}                                                                           
                                                                           
if ( typeof( repeat ) == "undefined" )
{
   String.prototype.repeat = function( n )                               
   { 
      var tmp = ""; 
      for ( var i = 0; i < n; i++ ) tmp += this;                                                                         
      return tmp;
   } 
} 

String.prototype.left = function( n, char )                               
{ 
   if ( typeof( char ) == "undefined" ) char = " ";  
   if ( this.length >= n ) return this.substr( 0, n );
   return this + char.repeat( n - this.length );
} 

String.prototype.right = function( n, char )                               
{
   if ( typeof( char ) == "undefined" ) char = " ";   
   if ( this.length > n ) return this.substr( this.length - n );
   return char.repeat( n - this.length ) + this;
}

Date.prototype.DateToDDMMYYYY = function()
{
   var YYYY = this.getFullYear();
   var MM   = ( this.getMonth() + 1 ).toString().right( 2, "0" );
   var DD   = this.getDate().toString().right( 2, "0" );

   var hh   = ( this.getHours()   ).toString().right( 2, "0" );
   var mm   = ( this.getMinutes() ).toString().right( 2, "0" );
   var ss   = ( this.getSeconds() ).toString().right( 2, "0" );

   return DD + "." + MM + "." + YYYY + " " + hh + ":" + mm + ":" + ss; 
}
//
// end-of-methods set
//

//
// Small class myStdReg() which provides methods to enumerate and 
// read Registry data of some types. This is simplified version
// of the https://github.com/sysprg-46/StdRegWrapperClass
// All the methods used in this class are derived from the
// correspondent methods of WMI-class root\default\StdRegProv 
//

function myStdReg()
{
   function StdRegFailed( method, objInParam, objOutParam )
   {
      var prefix    = null;
      var valuename = null;
      var msg = "";

      var GetMethod = ( method.indexOf( "Get" ) > - 1 );

      if ( GetMethod )
      {
         if ( method.indexOf( "StringValue" ) > - 1 ) prefix = "s";
         else prefix = "u";

         valuename = objInParam.sValueName;

         if ( typeof( valuename ) == "undefined" )
         {
            msg = "*** INTERNAL ERROR: Property sValuName"  +
            " is not set in input parameters for the method " + method;
            WScript.StdOut.WriteLine( msg );
            WScript.Quit();
         }
      }

      var root     = objInParam.hDefKey;
      var regPath  = objInParam.sSubKeyName;

      var msg = "StdRegProv method " + method + 
      " ReturnValue = " + objOutParam.ReturnValue +
      " for the path " + regPath;

      if ( GetMethod )
      {
         var value = objOutParam[ prefix + "Value" ];
         msg  += ", value name = " + valuename + 
                 " IS_NULL = " + ( value == null );
      }

      WScript.StdOut.WriteLine( msg );
      WScript.Quit();
   }

   var HKCR        = 0x80000000;// HKEY_CLASSES_ROOT
   var HKCU        = 0x80000001;// HKEY_CURRENT_USER
   var HKLM        = 0x80000002;// HKEY_LOCAL_MACHINE
   var HKU         = 0x80000003;// HKEY_USERS
   var HKCC        = 0x80000005;// HKEY_CURRENT_CONFIG
   
   var objLocator  = new ActiveXObject( "WbemScripting.SWbemLocator" ); 
   var objService  = objLocator.ConnectServer( ".", "root\\default" ); 
   var objReg      = objService.Get( "StdRegProv" ); 

   var objMethod   = null;
   var objInParam  = null;
   var objOutParam = null;
   var res         = null;

   var root = HKLM;
   
   this.RegReadStringValue = function ( root, regPath, valueName ) 
   { 
      var res = null, rc; 
      objMethod  = objReg.Methods_.Item( "GetStringValue" );
      objInParam = objMethod.InParameters.SpawnInstance_(); 
      objInParam.hDefKey     = root;     
      objInParam.sSubKeyName = regPath;  
      objInParam.sValueName  = valueName;
      
      res = null;
      objOutParam = objReg.ExecMethod_( objMethod.name, objInParam );   
      if( objOutParam.ReturnValue == 0 && objOutParam.sValue != null ) 
      { 
         res = objOutParam.sValue; 
      }
      return res; 
   }// End-of-RegReadStringValue

   this.RegReadDWORDValue = function ( root, regPath, valueName )
   { 
      var res = null; 
      objMethod  = objReg.Methods_.Item( "GetDWORDValue" );
      objInParam = objMethod.InParameters.SpawnInstance_(); 
      objInParam.hDefKey     = root;    
      objInParam.sSubKeyName = regPath; 
      objInParam.sValueName  = valueName;
      
      objOutParam = objReg.ExecMethod_( objMethod.name, objInParam );   

      if( objOutParam.ReturnValue == 0 && objOutParam.uValue != null ) 
      {  
         res = objOutParam.uValue; 
      }
      return res; 
   }// End-of-RegReadDWORDValue

   this.RegReadQWORDValue = function ( root, regPath, valueName )
   {     
      var res = null; 
      objMethod  = objReg.Methods_.Item( "GetQWORDValue" );
      objInParam = objMethod.InParameters.SpawnInstance_(); 
      objInParam.hDefKey     = root;    
      objInParam.sSubKeyName = regPath; 
      objInParam.sValueName  = valueName;
      
      objOutParam = objReg.ExecMethod_( objMethod.name, objInParam );   

      if( objOutParam.ReturnValue == 0 && objOutParam.uValue != null ) 
      {  
         res = objOutParam.uValue; 
      }
      return res; 
   }// End-of-RegReadQWORDValue


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

   //
   // this method converts registry REG_DWORD/REG_QWORD value
   // to JS Date object and returns it.
   // 
   this.RegDateToDate = function( regDate )
   {
      var MaxDWORD = 0x100000000;

      if ( regDate < MaxDWORD )// if RegDate is REG_DWORD
      {
         return new Date( 1000 * regDate );
      }
      // Input regDate is QWORD presenting FILETIME date value.
      var since1601 = 11644473600000; // since 01.01.1601 till 01.01.1970
      var MilliSeconds = Math.round( regDate / 10000 ) - since1601;
      return new Date( MilliSeconds );
   }
   this.InterpreteRegValueType = 
   {
       0: "REG_NONE",                       // No value type
       1: "REG_SZ",                         // Unicode nul terminated string
       2: "REG_EXPAND_SZ",                  // Unicode nul terminated string
                                            //    (with environment variable references)
       3: "REG_BINARY",                     // Free form binary
       4: "REG_DWORD",                      // 32-bit number
       5: "REG_DWORD_BIG_ENDIAN",           // 32-bit number
       6: "REG_LINK",                       // Symbolic Link (unicode)
       7: "REG_MULTI_SZ",                   // Multiple Unicode strings
       8: "REG_RESOURCE_LIST",              // Resource list in the resource map
       9: "REG_FULL_RESOURCE_DESCRIPTOR",   // Resource list in the hardware description
      10: "REG_RESOURCE_REQUIREMENTS_LIST", //
      11: "REG_QWORD"                       // 64-bit number
   };
}  
//
// End-of-myStdReg()-Class-code
//

/*
 Objects collection presentation function. Is called from the project tester code only. 
 Presents each object in one line. Output example follows.
 
BuildN   __ Install date ___  ___ Upgrade date __  ________ Product name __________  __Branch, UBR, ReleaseId__  Minutes
 9600   24.03.2015 22:09:18  31.03.2015 15:44:13  Windows 8.1 Pro                                               
10049 * 31.03.2015 16:25:14  23.04.2015 02:55:16  Windows 10 Pro Technical Preview                                 41
10061   23.04.2015 03:43:12  29.04.2015 06:35:12  Windows 10 Pro Technical Preview                                 48
10074   29.04.2015 07:24:42  21.05.2015 11:31:06  Windows 10 Pro Insider Preview                                   50
10122   21.05.2015 12:22:52  28.05.2015 21:12:36  Windows 10 Pro Insider Preview    fbl_impressive     0           52
10125   28.05.2015 22:03:15  30.05.2015 19:10:21  Windows 10 Pro Insider Preview    fbl_impressive     0           51
10130   30.05.2015 20:15:42  01.07.2015 11:16:36  Windows 10 Pro Insider Preview    fbl_impressive     0           65
-------------------------------------------------------------------------------------------------------
*/
function BriefShow( arr )
{
   var header = "BuildN   __ Install date ___  ___ Upgrade date __  ________ Product name __________  __Branch, UBR, ReleaseId__" +
   "  Minutes";

   var bObj, msg = "", tmp;
   var maxL = "Windows 10 Pro Technical Preview".length;
   var timestampL = "2015.05.28 21:13:37".length;
   var branchL = "fbl_impressive".length;
   WScript.StdOut.WriteLine( header + "\n" );

   for ( var i = 0; i < arr.length; i++ )
   {
      bObj = arr[ i ];
      msg += bObj.BuildNumber.toString().right( 5 ) + " " + ( bObj.InstallOnTheFlight ? "* " : "  " );
      msg += bObj.InstallDateString + "  ";
      msg += ( i == arr.length - 1 ? " ".left( timestampL ) : bObj.UpgradeDateString ) + "  "; 
      msg += bObj.ProductName.left( maxL ) + "  ";
      tmp = ( bObj.BuildBranch ? bObj.BuildBranch.left( branchL ) : "  ".left( branchL ) );
      msg += tmp + "  ";
      tmp = ( typeof( bObj.UBR ) == "number" ? bObj.UBR.toString().right( 4 ) : " ".left( 4 ) );
      msg += tmp + "  ";
      msg += ( bObj.ReleaseId ? bObj.ReleaseId : " ".left( 4 ) ) + "  ";
      if ( typeof( bObj.InstallDuration ) == "string" )
      {
         msg += bObj.InstallDuration.right( 5 );
      }
      WScript.StdOut.WriteLine( msg );
      msg = "";
   }
   WScript.StdOut.WriteLine( "\n" );
   WScript.StdOut.WriteLine( "*** History contains " + arr.length + " upgrades\n" );
   WScript.StdOut.WriteLine( '*** Builds installed "On-the-Flight" are marked by the star\n' );
   WScript.StdOut.WriteLine( '*** Column "Minutes" contains installation duration in minutes' );
   WScript.StdOut.WriteLine( "\n\n" );
   return false;
}

/*
 Objects collection presentation function.
 Each property of each of the objects is printed in one line.
 Output example for one of the objects follows

------------------------------------------------------------
BuildNumber        11082
BuildLabEx         11082.1000.amd64fre.rs1_release.151210-2021
RegPath            SYSTEM\Setup\Source OS (Updated on 1/14/2016 10:47:31)
InstallDateString  17.12.2015 00:31:03
UpgradeDateString  14.01.2016 10:47:31
InstallDuration    52
ExplainDuration    52 minutes = (1450297863-1450294758)/60
InstallDate        1450297863
UpgradeDate        1452754051
InstallTime        130947714635122379
ProductName        Windows 10 Pro Insider Preview
ProductId          00330-80000-00000-AA909
BuildBranch        rs1_release
ReleaseId          1511
UBR                1000
BuildWasReleased   17.12.2015
InstallOnTheFlight true
-------------------------------------------------------------
*/
function DebugShow( arr )
{
   var bObj, msg = "", val, valx;
   var maxPL = "InstallOnTheFlight".length;

   for ( var i = 0; i < arr.length; i++ )
   { 
      bObj = arr[ i ]; 
      for ( var p in bObj )
      {
         WScript.StdOut.WriteLine( p.left( maxPL ) + " " + bObj[ p ] );
      }
      WScript.StdOut.WriteLine( "-------------------\n" );
   }

   return false;
}

  
/*
-------------------------------------------------------------------------------------
Main project function: it builds objects array. 
-------------------------------------------------------------------------------------
This function will create an object for each of the Builds detected in the Registry. 
Objects will be sorted by the build numbers. Each object has some primary properties
and some secondary. Primary properties are those which values were set on base of the 
data read from the Registry. Secondary propierties are those which were derived from 
the primary properties.
-------------------------------------------------------------------------------------
*/  
function BuildsHistory()
{
   var BuildN2ReleaseDate =
   {
       7601:  "15.03.2011", // SP1
       9600:  "27.08.2013", // RTM: August 27, 2013, GA: October 17, 2013
       9841:  "01.10.2014",
       9860:  "21.10.2014",
       9879:  "12.11.2014",
       9926:  "23.01.2015",
      10041:  "18.03.2015",
      10049:  "31.03.2015",
      10061:  "22.04.2015",
      10074:  "26.04.2015",
      10122:  "20.05.2015",
      10125:  "25.05.2015",
      10130:  "29.05.2015",
      10134:  "05.06.2015",
      10135:  "05.06.2015",
      10147:  "18.06.2015",
      10158:  "30.06.2015",
      10159:  "30.06.2015",
      10162:  "02.07.2015",
      10166:  "09.07.2015",
      10240:  "15.07.2015", // Windows 10 RTM
      10525:  "08.08.2015",
      10532:  "28.08.2015",
      10547:  "18.09.2015",
      10565:  "12.10.2015",
      10576:  "29.10.2015",
      10586:  "05.11.2015", // th1
      11082:  "17.12.2015",
      11099:  "13.01.2016",
      11102:  "20.01.2016",
      14251:  "27.01.2015",
      14257:  "04.02.2016",
      14267:  "18.02.2016",
      14271:  "24.02.2016",
      14279:  "04.03.2016",
      14291:  "18.03.2016",
      14295:  "25.03.2016",
      14316:  "06.04.2016",
      14328:  "22.04.2016",
      14332:  "27.04.2016",
      14342:  "11.05.2016",
      14352:  "27.05.2016",
      14361:  "08.06.2016",
      14366:  "15.06.2016",
      14367:  "17.06.2016",
      14371:  "22.06.2016",
      14372:  "24.06.2016",
      14376:  "29.06.2016",
      14379:  "01.07.2016",
      14383:  "08.07.2016",
      14385:  "10.07.2016",
      14388:  "13.07.2016",
      14390:  "15.07.2016",
      14393:  "19.07.2016", //
      14901:  "12.08.2016",
      14905:  "17.08.2016",
      14915:  "31.08.2016",
      14926:  "14.09.2016",
      14931:  "22.09.2016",
      14936:  "28.09.2016",
      14942:  "07.10.2016",
      14946:  "13.10.2016",
      14951:  "20.10.2016"        
   };

   ///////////////////////////////////////////////////////////////////////////////
   var R500 = false; // Developer hint: is set true when run at developer Thinkpad
   ///////////////////////////////////////////////////////////////////////////////

   //
   // Input:
   // regObj - myStdReg class object created by the caller
   // path   - Registry path which contains necessary values
   // bObj   - Build object created by the caller
   // Function gets necessary data from the Registry using
   // methods of the class myStdReg and set bObj properties.
   // Updated bObj is returned to the caller. 
   // Function is called some times from the main function
   // BuildsHistory() and returns updated object.
   // 
   function SetBuildObjectProperties( regObj, path, bObj )
   {
      var HKLM = 0x80000002;// HKEY_LOCAL_MACHINE
      var res, props;

      bObj.BuildLabEx  = regObj.RegReadStringValue( HKLM, path, "BuildLabEx" );
      bObj.BuildNumber = parseInt( bObj.BuildLabEx.split( "." )[ 0 ] );
      bObj.InstallDate = regObj.RegReadDWORDValue( HKLM, path, "InstallDate" );
      bObj.InstallDateString = regObj.RegDateToDate( bObj.InstallDate ).DateToDDMMYYYY(); 
      if ( bObj.BuildNumber > 9600 )
      {
         bObj.InstallTime = regObj.RegReadQWORDValue( HKLM, path, "InstallTime" );
      }
      bObj.ProductName = regObj.RegReadStringValue( HKLM, path, "ProductName" );
      bObj.ProductId   = regObj.RegReadStringValue( HKLM, path, "ProductId" );
      res = regObj.RegReadStringValue( HKLM, path, "BuildBranch" );   
      if ( typeof( res ) == "string" ) { bObj.BuildBranch = res; } 
      res              = regObj.RegReadStringValue( HKLM, path, "ReleaseId" );
      if ( typeof( res ) == "string" ) { bObj.ReleaseId = res; }
      res = regObj.RegReadDWORDValue(  HKLM, path, "UBR" ); 
      if ( typeof( res ) == "number" ) { bObj.UBR = res; } 
  
      return bObj;
   }

   // 
   // Function counts duration of the Windows 10 build installation on base of 
   // the two Build objects.
   // New - current  Build object in the sorted array of the builds objects
   // Old - previous Build object in the sorted array of the builds objects
   // New.InstallDate - indicates End of the installation of the current build
   // Old.UpgradeDate - indicates Start of the installation of the current build 
   // 
   function CountDuration( New, Old )
   {
      var Seconds = New.InstallDate - Old.UpgradeDate;
      var Minutes = Math.round( Seconds / 60 );
      return Minutes.toString();
   }

   function ReorderProperties( obj )
   {
      var tmp = {};
      tmp.BuildNumber = obj.BuildNumber;
      tmp.BuildLabEx  = obj.BuildLabEx;
      tmp.RegPath  = obj.RegPath;
      tmp.InstallDateString = obj.InstallDateString;
      if ( obj.UpgradeDateString != null )
      {
         tmp.UpgradeDateString = obj.UpgradeDateString;
      }
      if ( typeof( obj.InstallDuration ) == "string" )
      {
         tmp.InstallDuration = obj.InstallDuration;
         tmp.ExplainDuration = obj.ExplainDuration;
      }
      tmp.InstallDate = obj.InstallDate;
      if ( obj.UpgradeDate != null )
      {
         tmp.UpgradeDate = obj.UpgradeDate;
      }      
      if ( tmp.BuildNumber > 9600 )
      {
         tmp.InstallTime = obj.InstallTime;
      }
      tmp.ProductName  = obj.ProductName;
      tmp.ProductId  = obj.ProductId;
      if ( tmp.BuildNumber >= 10122 )
      {
         tmp.BuildBranch = obj.BuildBranch;
      }
      if ( tmp.BuildNumber >= 10565 )
      {
         tmp.ReleaseId = obj.ReleaseId;
      }
      if ( tmp.BuildNumber >= 10122 )
      {
         tmp.UBR = obj.UBR;
      }
      tmp.BuildWasReleased = obj.BuildWasReleased;
      tmp.InstallOnTheFlight = obj.InstallOnTheFlight;
      return tmp;
   }

   var myRegObj = new myStdReg();// create service object to work with Registry

   var HKLM = 0x80000002;// HKEY_LOCAL_MACHINE

   var bObj = {}, builds = [], rc, res;

   ////////////////////////////////////////////////////////////////////////////////////////////////////////////
   //
   // Verify if we are running at developer computer:
   // in this case will need to correct time zone for the string dates
   // less then 14.01.2016 as all previous dates were saved in the 
   // time zone which differs current one.
   //
   var path = "HARDWARE\\DESCRIPTION\\System\\MultifunctionAdapter\\0\\DiskController\\0\\DiskPeripheral\\0";
   var MyId = myRegObj.RegReadStringValue( HKLM, path, "Identifier" );

   R500 = ( MyId == "9f9c4a4c-80d01f2a-A" ); // is true until developer SSD is alive... 

   if ( R500 )
   {
      WScript.StdErr.WriteLine( "*** Running at developer computer ***" );
   }
   ////////////////////////////////////////////////////////////////////////////////////////////////////////////

   //
   // Step 1, we will set some of the current Build object properties. 
   // HKLM\SYSTEM\Setup, value with the name CloneTag contains date of uninstallation
   // of the previous build and build number of the previous build. Though it seems
   // that time stamp of the previous build uninstallation is also timestamp for the 
   // current build installation, we will use another source to get precise value 
   // of this date together with some more properties of the current build object.
   //

   var regPath1 = "SYSTEM\\Setup";
   var regPath2 = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion"; 

   bObj = SetBuildObjectProperties( myRegObj, regPath2, {} );
   bObj.RegPath = regPath2;
   builds[ 0 ]  = bObj;

   //
   // Step 2, we will enumerate all builds present in HKLM\SYSTEM\Setup. 
   // They are pesented by subkeys with a names like 
   // "Source OS (Updated on 3/31/2015 14:44:13)". Thus we will enumerate
   // subkeys of HKLM\SYSTEM\Setup, then filer those subkeys which
   // corresponds to the builds and create objects for them.
   //

   var key = regPath1, path, p, needKey = "Source OS (Updated on ";
   var needKeyL = needKey.length, tmp1, tmp2, tmp3, rc;
 
   res = myRegObj.RegEnumSubKeys( HKLM, regPath1 );

   for ( var i = 0; i < res.length; i++ )  // In this loop we filter out 
   {                                       // unnecessary subkeys. 
      p = res[ i ].indexOf( needKey );     // Try to detect if we need subkey
      if ( p > - 1 )                       // Yes, subkey presents Build and
      {                                    // we will create Build object
         bObj = new Object;                // and set object properties

         bObj.RegPath = regPath1 + "\\" + res[ i ];

         p += needKeyL;

         // get date field from subkey and save it in tmp1
         tmp1 = res[ i ].substr( p, res[ i ].indexOf( ")", p ) - 1 );
         tmp1 = tmp1.substr( 0, tmp1.length - 1 ).strip();

         /////////////////////////////////////////////////////////////////////
         if ( R500 //&& res[ i ].BuildNumber < 10586 )
         &&   tmp1.indexOf( "2015"       )  > - 1 
         &&   tmp1.indexOf( "12/16/2015" ) == - 1 )
         {
            WScript.StdErr.WriteLine( tmp1 + "\n" ); 
         /*
            This is "What-For" R500 variable was defined: set correct
            time zone for the Upgrade Dates presented by the Strings 
            like "Source OS (Updated on 7/16/2015 09:58:26)"
         */
            tmp1 += " UTC+03"; // set correct time zone
         }
         /////////////////////////////////////////////////////////////////////

         tmp3 = new Date( tmp1 );
         bObj.UpgradeDate = Math.round( tmp3.getTime() / 1000 );
         bObj.UpgradeDateString = tmp3.DateToDDMMYYYY();

         path = key + "\\" + res[ i ];  // prepare path for reading values
         bObj = SetBuildObjectProperties( myRegObj, path, bObj );
         builds[ builds.length ]= bObj;     
      }
   }
 
   //
   // Step 3, define Sort function for the array of the builds 
   // and sort that array by the key = ( BuildNumber, InstallDate )
   //
   builds.sort( function( a, b )
   {
      if ( a.BuildNumber < b.BuildNumber ) return 1;
      if ( a.BuildNumber > b.BuildNumber ) return -1;
      if ( a.InstallDate < b.InstallDate ) return 1;
      if ( a.InstallDate > b.InstallDate ) return -1; 
      return 0;
   } );

   builds.sort(); // 

   //
   // Step 4, now we can count Installation duration as a value
   // of ( previous object UpgradeDate - current object InstallDate )
   //
   for ( var i = 0; i < builds.length; i++ )
   {
      tmp1 = BuildN2ReleaseDate[ builds[ i ].BuildNumber ];
      builds[ i ].BuildWasReleased = tmp1;
      tmp2 = builds[ i ].InstallDateString.split( " " )[ 0 ];
      builds[ i ].InstallOnTheFlight = ( tmp1 == tmp2 );

      if ( i > 0 )
      {
         builds[ i ].InstallDuration = CountDuration( builds[ i ], builds[ i - 1 ] ); 
         builds[ i ].ExplainDuration = builds[ i ].InstallDuration + " minutes = (" + 
         builds[ i ].InstallDate.toString() + "-" + builds[ i - 1 ].UpgradeDate.toString() + ")/60";
      }

   //
   // Step 5, reorder objects properties
   //
      builds[ i ] = ReorderProperties( builds[ i ] );
   }
                                        // all Builds presentaion objects were
   return builds;                       // created. Return to the caller an array
}                                       // of the objects created.

//--------------------------- Project testing code -------------------                                                                         
var rc;
var arr = BuildsHistory(); // create an array of Builds objects

rc = BriefShow( arr );     // get brief presentation
rc = DebugShow( arr );     // get full presentation

var objShell = new ActiveXObject( "WScript.Shell" );
objShell.run( "notepad.exe BuildsHistory.log" );
WScript.Quit(); 

//--End-of-File-BuildsHistory------------------------------------------ 

