//--------------------------------------------------------------------- 
// Project: CompInfo                                                    
// Version: 1.0                                                         
// FileId:  e:\Diary\MYCOMP~1\R500\CompInfo                             
// When:    3 Nov 2016,  Thursday,  11:41:52                              
// Who:     Oleg Kulikov, sysprg@live.ru                                
//--------------------------------------------------------------------- 
// What for:  this is FAD prototype code to collect and present hardware 
// information.                                                         
//--------------------------------------------------------------------- 

'use strict';

var 
Version = "0.1 03-11-2016";
Version = "0.2 04-11-2016";
Version = "0.3 07-11-2016";
Version = "0.4 08-11-2016";

var RAMOnly    = false;
var Globals    = {};
Globals.Select = "*"; // use all classes provided select parameter is absent

if ( WScript.Arguments.Named.Length > 0
&&   typeof( WScript.Arguments.Named.Item( "select" ) ) != null )
{
   Globals.Filter = WScript.Arguments.Named.Item( "select" );
   var work   = WScript.Arguments.Named.Item( "select" ).toUpperCase().split(",");
   var select = {};
   var translate = 
   {
      "PLATFORM": "ComputerSystem", "CPU": "Processor", "RAM": "PhysicalMemory",
      "VIDEO": "VideoController", "DISKS": "DiskDrive", "SOUND": "SoundDevice",
      "BATTERY": "PortableBattery"
   };
   for ( var i = 0; i < work.length; i++ )
   {
      var p = work[ i ];
      if ( "PLATFORM,CPU,RAM,DISKS,VIDEO,SOUND,BATTERY".indexOf( p ) > - 1 )
      {
         select[ "Win32_" + translate[ p ] ] = true;
      }
   }
   Globals.Select = select;
   if ( work.length == 1 && Globals.Select.Win32_PhysicalMemory )
   {
      RAMOnly = true;
   }
} 
  
var scriptn = WScript.ScriptName.split( "." )[0].toLowerCase();

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
String.prototype.strip = function()                                         
{                                                                           
   return this.replace( /^(\s*)/, "" ).replace( /(\s*)$/, "" );               
} 

function Explain( component, p, v )
{ 
// https://msdn.microsoft.com/en-us/library/aa394347(v=vs.85).aspx - FormFactor, MemoryType
// P.A: value for MemoryType[ 23 ] - is not defined !!!
   var Translation = 
   {
      "PhysicalMemory": 
      {   
         "FormFactor": // 0..23
         [  
            "Unknown", "Other", "SIP", "DIP", "ZIP", "SOJ", "Proprietary", "SIMM", 
            "DIMM", "TSOP", "PGA", "RIMM", "SODIMM", "SRIMM", "SMD", "SSMP", "QFP",
            "TQFP", "SOIC", "LCC", "PLCC", "BGA", "FPBGA", "LGA"
         ],
         "MemoryType": // 0..25
         [
            "Unknown", "Other", "DRAM", "Synchronous DRAM", "Cache DRAM", 
            "EDO", "EDRAM", "VRAM", "SRAM", "RAM", "ROM", "Flash", "EEPROM",
            "FEPROM", "EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM", "RDRAM",
            "DDR", "DDR2", "DDR2 FB-DIMM", "Undefined", "DDR3", "FBD2"
         ]  
      },// end-of-"PhysicalMemory"
      "BIOS":
      {
         "BiosCharacteristics": // 0..39
         [
            "Reserved","Reserved","Unknown","BIOS Characteristics Not Supported",
            "ISA is supported","MCA is supported","EISA is supported","PCI is supported",
            "PC Card (PCMCIA) is supported","Plug and Play is supported","APM is supported",
            "BIOS is Upgradeable (Flash)","BIOS shadowing is allowed","VL-VESA is supported", 
            "ESCD support is available","Boot from CD is supported","Selectable Boot is supported", 
            "BIOS ROM is socketed","Boot From PC Card (PCMCIA) is supported", 
            "EDD (Enhanced Disk Drive) Specification is supported", 
            "Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5\", 1k Bytes/Sector, 360 RPM) is supported", 
            "Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5\", 360 RPM) is supported",
            "Int 13h - 5.25\"( / 360 KB Floppy Services are supported",
            "Int 13h - 5.25\"( /1.2MB Floppy Services are supported", 
            "Int 13h - 3.5\"( / 720 KB Floppy Services are supported", 
            "Int 13h - 3.5\"( / 2.88 MB Floppy Services are supported", 
            "Int 5h, Print Screen Service is supported","Int 9h, 8042 Keyboard services are supported", 
            "Int 14h, Serial Services are supported","Int 17h, printer services are supported", 
            "Int 10h, CGA/Mono Video Services are supported","NEC PC-98","ACPI supported", 
            "USB Legacy is supported","AGP is supported","I2O boot is supported", 
            "LS-120 boot is supported","ATAPI ZIP Drive boot is supported", 
            "1394 boot is supported","Smart Battery supported" 
         ]
      },
      "PortableBattery":
      {
         "Chemistry":
         [ 
            "Notapplicable","Other","Unknown","Lead Acid","Nickel Cadmium","Nickel Metal Hydride",
            "Lithium-ion","Zinc air","Lithium Polymer"
         ],
         "BatteryStatus":
         [
            "The battery is discharging.","The system has access to AC so no battery is being discharged.",
            "Fully Charged","Low","Critical","Charging","Charging and High","Charging and Low",
            "Charging and Critical","Undefined","Partially Charged"
         ]
      }         
   };

   function IntoBra( v )
   {
      var t = v;
      if ( typeof( v ) == "number" ) { t = t.toString(); }
      return " (" + t + ")";
   }
  
   function BytesToString( v )
   {
      if ( v < 1024 ) return v.toString() + " bytes";
      v /= 1024;
      if ( v < 1024 ) return v.toString() + " KB";
      v /= 1024;
      if ( v < 1024 ) return v.toString() + " MB";
      v /= 1024;
      if ( v < 1024 ) return ( v.toFixed( 2 ) ).toString() + " GB";
      v /= 1024;
      if ( v < 1024 ) return ( v.toFixed( 2 ) ).toString() + " TB";
      v /= 1024;
      return v.toString() + " PB";
   } 
   switch ( component )
   {  
      case "Win32_PortableBattery":
      {
         switch ( p )
         {
            case "BatteryStatus": 
            {
               return (v >= 1 && v < 11) ? v + IntoBra( Translation.PortableBattery.BatteryStatus[ v ] ) :
                                            v + " (Unknown)";
            }
            case "Chemistry": 
            {
               return (v >= 1 && v < 9 ) ? v + IntoBra( Translation.PortableBattery.Chemistry[ v ] ) :
                                            v + " (Unknown)";
            }
         }
      }
      case "Win32_BIOS":
      {
         if ( p == "BiosCharacteristics" )
         {
            var arr = v.split( "," );
            var lines = [ v ], cv, iv;
            for ( var j = 0; j < arr.length; j++ )
            {
               iv = parseInt( arr[ j ] );
               cv = arr[ j ].left( 2 );
               lines[ lines.length ] = iv < 40 ? cv + IntoBra( Translation.BIOS.BiosCharacteristics[ iv ] ) :
                                       iv < 48 ? cv + " (Reserved for BIOS vendor)"   :
                                       iv < 63 ? cv + " (Reserved for system vendor)" : cv + " (Unknown)";                                      
            }
            return lines;            
         } 
      }                 
      case "Win32_PhysicalMemory":
      {
         switch ( p )
         {
            case "FormFactor": 
            {
               return (v >= 0 && v < 23) ? v + IntoBra( Translation.PhysicalMemory.FormFactor[ v ] ) :
                                            v + " (Unknown)";
            }
            case "MemoryType": 
            {
               return (v >= 0 && v < 25) ? v + IntoBra( Translation.PhysicalMemory.MemoryType[ v ] ) :
                                            v + " (Unknown)";
            }
            case "Capacity": 
            { 
               return BytesToString( v ); 
            }
            case "Manufacturer":
            { 
               if ( v.strip() == "80AD" ){ return v.strip() + " (Hynix)";} 
            }
            case "TypeDetail":
            { 
               switch( v )
               {                     
                  case    1: { return v + " (Reserved)"     ; }
                  case    2: { return v + " (Other)"        ; }
                  case    4: { return v + " (Unknown)"      ; }
                  case    8: { return v + " (Fast-Paged)"   ; }
                  case   16: { return v + " (Static column)"; }
                  case   32: { return v + " (Pseudo-static)"; }
                  case   64: { return v + " (RAMBUS)"       ; }
                  case  128: { return v + " (Synchronous)"  ; }
                  case  256: { return v + " (CMOS)"         ; }
                  case  512: { return v + " (EDO)"          ; }
                  case 1024: { return v + " (Window DRAM)"  ; }
                  case 2048: { return v + " (Cache DRAM)"   ; }
                  case 4096: { return v + " (Non-volatile)" ; }
               }
            }// end-of-case "TypeDetail":

         }// end-of-case switch ( p )
      }// end-of-case Win32_PhysicalMemory":
      case "Win32_VideoController":
      {
         if ( p == "AdapterRAM" )
         { 
            return BytesToString( v ); 
         }                  
      }
      case "Win32_DiskDrive":
      {
         if ( p == "Size" )
         { 
            return BytesToString( v ); 
         }    
      } 
      case "Win32_Processor":
      {
         switch( p )
         {
            case "L2CacheSize":    { return BytesToString( v )     ; }
            case "LoadPercentage": { return v.toString() + " %"    ; }
            case "ExtClock":       { return v.toString() + " MGhz" ; }
            case "AddressWidth":   { return v.toString() + " bits" ; }
            case "DataWidth":      { return v.toString() + " bits" ; }
            case "CurrentClockSpeed": return ( v / 1000 ).toFixed( 2 ).toString() + " Ghz";  
            case "Architecture":
            { 
               switch( v )
               {                     
                  case 0: { return v + " (x86)"    ; }
                  case 1: { return v + " (MIPS)"   ; }
                  case 2: { return v + " (Alpha)"  ; }
                  case 4: { return v + " (PowerPC)"; }
                  case 5: { return v + " (ARM)"    ; }
                  case 6: { return v + " (ia64)"   ; }
                  case 9: { return v + " (x64)"    ; }
               } 
            } 
            case "ProcessorType":
            { 
               switch( v )
               {                     
                  case 1: { return v + " (Other)"             ; }
                  case 2: { return v + " (Unknown)"           ; }
                  case 3: { return v + " (Central Processor)" ; }
                  case 4: { return v + " (Math Processor)"    ; }
                  case 5: { return v + " (DSP Processor)"     ; }
                  case 6: { return v + " (iVideo Processor"   ; }
               } 
            }         
         }// end-of-switch( p )                
      }// end-of case "Win32_Processor": 
   }// end-of-switch ( component )
   if ( typeof( v ) == "string" ) return v.strip().replace( /\s\s*/g, " " );
   return v;
}   
 
function ShowProps( CompObj ) 
{
   var Log      = [];
   var Props    = [];
   var PropsObj = {};
   var MaxL     = 1 + "SecondLevelAddressTranslationExtensions".length;
   var dashes   = "-".left( 40, "-" );
   var v;
   var msg = ( new Date() ).toString() + ", " + scriptn + ".js, " +
             "version " + Version + " runs at\n" + Globals.Platform;
   Log[ Log.length ] = msg + "\n"; 
   if ( Globals.Filter )
   {
      msg = "Output was filtered accordingly selection passed in command line: ";     
      Log[ Log.length ] = msg  + "select:" + Globals.Filter + "\n";
   }      
   for ( var component in CompObj ) 
   {
      if ( RAMOnly && component != "Win32_PhysicalMemory" ) continue;
      v = component.split( "_" )[ 1 ];
      Log[ Log.length ] = dashes + v + dashes.substr( 0, 40 - v.length );
      Props = CompObj[ component ];
      for ( var i = 0; i < Props.length; i++ )
      {
         PropsObj = Props[ i ];
         for ( var p in PropsObj )
         {
            v = PropsObj[ p ];
            if ( typeof( v ) != "object" )
            {
               Log[ Log.length ] = ( p + ":" ).left( MaxL ) + v;
            }
            else if ( v.length > 1 )
            {
               Log[ Log.length ] = ( p + ":" ).left( MaxL ) + v[0];
               for( var j = 1; j < v.length; j++ )
               {
                  Log[ Log.length ] = "   ".left( MaxL ) + v[j];   
               }
            }
         }
         if ( Props.length > 1 && i < Props.length - 1 )
         {
            Log[ Log.length ] = "\n";
         }
      }
   }
   return Log;
}

function ComputerInfo()
{
   var Win32Service = GetObject( "winmgmts:\\\\.\\root\\CIMV2" );
   var Classes = {};
   var All = 
   {
      "Win32_ComputerSystem":        "Manufacturer,Model,SystemFamily,SystemType,TotalPhysicalMemory,UserName",
      
      "Win32_ComputerSystemProduct": "Description,IdentifyingNumber,Name,SKUNumber,Vendor,NumberOfCores,",
      
      "Win32_BaseBoard":             "Description,CreationClassName,HostingBoard,dt-InstallDate,Manufacturer," +
                                     "Model,Name,OtherIdentifyingInfo,PartNumber,Product,SKU,Version",
                                     
      "Win32_BIOS":                  "Description,ta-BiosCharacteristics,ta-BIOSVersion,BuildNumber,"         + 
                                     "CodeSet,CurrentLanguage,EmbeddedControllerMajorVersion,"                +
                                     "EmbeddedControllerMinorVersion,IdentificationCode,"                     +
                                     "InstallableLanguages,dt-InstallDate,ta-ListOfLanguages,"                +
                                     "dt-ReleaseDate,SerialNumber,SMBIOSBIOSVersion,"                         +          
                                     "SMBIOSMajorVersion,SMBIOSMinorVersion,SystemBiosMajorVersion,"          +
                                     "SystemBiosMinorVersion,Version", 
                                     
      "Win32_Processor":             "Manufacturer,Description,AddressWidth,Architecture,CurrentClockSpeed,"  +
                                     "CurrentVoltage," +
                                     "DataWidth,ExtClock,L2CacheSize,L2CacheSpeed,L3CacheSize,L3CacheSpeed,"  +
                                     "Level,LoadPercentage,Name,NumberOfCores,NumberOfLogicalProcessors,"     +
                                     "ProcessorId,ProcessorType,PowerManagementSupported,"                    +
                                     "SecondLevelAddressTranslationExtensions,VirtualizationFirmwareEnabled," +
                                     "VMMonitorModeExtensions,SerialNumber,Version",
                                     
      "Win32_PhysicalMemory":        "Description,BankLabel,PositionInRow,Capacity,ConfiguredClockSpeed,"     +
                                     "ConfiguredVoltage,DataWidth,FormFactor,"                                +
                                     "InterleaveDataDepth,InterleavePosition,Manufacturer,PartNumber,"        +
                                     "SerialNumber,SKU,MaxVoltage,MinVoltage,MemoryType,Speed,TypeDetail", 
                                     
      "Win32_VideoController":       "Description,AdapterCompatibility,AdapterDACType,AdapterRAM," +
                                     "dt-DriverDate,DriverVersion,VideoModeDescription", 
                                     
      "Win32_DesktopMonitor":        "Description,MonitorManufacturer",
      
      "Win32_SoundDevice":           "Description,Manufacturer,DMABufferSize,ProductName",
      
      "Win32_DiskDrive":             "Caption,FirmwareRevision,InterfaceType,DeviceID,SerialNumber,Size,MediaType",
      
      "Win32_PortableBattery":       "Description,Availability,BatteryStatus,CapacityMultiplier,Chemistry," +
                                     "DesignCapacity,FullChargeCapacity,EstimatedChargeRemaining,"          +
                                     "ExpectedBatteryLife,ExpectedLife,InstallDate,ManufactureDate,"        +
                                     "Manufacturer,Location,Status,TimeOnBattery,TimeToFullCharge",
      
      "Win32_WinSAT":                "CPUScore,D3DScore,DiskScore,GraphicsScore,MemoryScore,TimeTaken," +
                                     "WinSATAssessmentState,WinSPRLevel"         
   };
   
   var tmp = Globals.Select;
   if ( typeof ( tmp ) == "string" ) Classes = All;
   else if ( typeof( tmp ) == "object" )
   {
      Classes.Win32_ComputerSystem = All.Win32_ComputerSystem; // always is necessary
      for ( var p in tmp )
      {
         Classes[ p ] = All[ p ];
      }
   }

   function WMIDateStringToDate( dtmDate )
   {
      if ( dtmDate == null )
      {
         return "null date";
      }
      var strDateTime;
      if ( dtmDate.substr( 4, 1 ) == 0 )
      {
         strDateTime = dtmDate.substr( 5, 1 ) + "/";
      }
      else
      {
         strDateTime = dtmDate.substr( 4, 2 ) + "/";
      }
      if (dtmDate.substr(  6, 1) == 0 )
      {
         strDateTime = strDateTime + dtmDate.substr( 7, 1 ) + "/";
      }
      else
      {
         strDateTime = strDateTime + dtmDate.substr( 6, 2 ) + "/";
      }
      return strDateTime + 
      dtmDate.substr(  0, 4 ) + " " +
      dtmDate.substr(  8, 2 ) + ":" +
      dtmDate.substr( 10, 2 ) + ":" +
      dtmDate.substr( 12, 2 );
   }

   var wbemFlagReturnImmediately = 0x10;
   var wbemFlagForwardOnly       = 0x20;
   var flag                      = wbemFlagReturnImmediately | wbemFlagForwardOnly;
   var comp                      = {};
   //
   // failed to run SQL request with my Select for BIOS
   // besides it seems that request runs faster for "*"   
   //
   for ( var CN in Classes )
   {
      var Select    = Classes[ CN ];
      var colItems  = Win32Service.ExecQuery( "Select * From " + CN, "WQL", flag );
      var EnumObj   = new Enumerator( colItems );
      var ret       = []; 
      var PropsList = Select.split( "," );  
                   
      for (; ! EnumObj.atEnd(); EnumObj.moveNext() ) 
      {
         PropsObj = EnumObj.item();
         var res = {};
         for ( var i = 0; i < PropsList.length; i++ )
         {
            var p = PropsList[ i ];
            if ( p.substr( 0, 3 ) == "ta-" )
            {
               p = p.substr( 3 );
               var v = null;
               try      
               { 
                  v = PropsObj[ p ].toArray().join(",");
                  res[ p ] = Explain( CN, p, v );           
               }
               catch(e) {}
               continue;
            }
            if ( p.substr( 0, 3 ) == "dt-" )
            {
               p = p.substr( 3 );
               var t = WMIDateStringToDate( PropsObj[ p ] );
               if ( t == "null date" ) continue;
               res[ p ] = t;
               continue;
            }
            if ( PropsObj[ p ] == null ) continue;
            if ( typeof( PropsObj[ p ] ) == "string" 
            &&   PropsObj[ p ].strip() == "" ) continue;
            res[ p ] = Explain( CN, p, PropsObj[ p ] );
         }// End-of-for ( var i = 0; i < PropsList.length; i++ )
         ret[ ret.length ] = res;
      }// End-of-for (; ! EnumObj.atEnd(); EnumObj.moveNext() )
      comp[ CN ] = ret;
   }// End-of-for ( var CN in Classes )
   //
   // Set ComputerSystem.TotalPhysicalMemory to the sum of RAM-chips capacity 
   //   
   var arr   = comp.Win32_PhysicalMemory;
   var Total = 0;
   for ( var i = 0; i < arr.length; i++ )
   {
      Total += parseInt( ( arr[ i ].Capacity ).split( " " )[0] );
   }
   ( comp.Win32_ComputerSystem )[0].TotalPhysicalMemory = Total.toString() + " GB";
   Globals.Platform = ( comp.Win32_ComputerSystem )[0].Manufacturer + "_" +
                      ( comp.Win32_ComputerSystem )[0].SystemFamily + "_" +
                      ( comp.Win32_ComputerSystem )[0].Model;
   Globals.Platform = Globals.Platform.replace( /\s/g, "_" ).replace( /__*/g, "_" );
   Globals.Platform = Globals.Platform.replace( "_undefined", "" );
   return comp;
}// End-of-ComputerInfo

////////////////////////////////////////////////////////////////////////////////////
var fso  = new ActiveXObject( "Scripting.FileSystemObject" );
var path = fso.GetAbsolutePathName( "." );
var Log  = ShowProps( ComputerInfo() );
WScript.Echo( Log.join( "\n" ) );
WScript.Quit();

//--End-of-File-CompInfo----------------------------------------------- 


