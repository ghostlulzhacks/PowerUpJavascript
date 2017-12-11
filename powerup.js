
function funcLocalAdminGroup()
{
	var admin = 0;
	var WSHnetwork = new ActiveXObject("WScript.Network");
	var objAdmin = GetObject("WinNT://" + WSHnetwork.ComputerName + "/Administrators");
	var colMembers = objAdmin.Members();
	var items = new Enumerator(colMembers);
	while (!items.atEnd())
	{
			objmember = items.item();
			if(WSHnetwork.UserName.toLowerCase() == objmember.Name.toLowerCase())
			{
				admin = 1;
			}
			items.moveNext();
	}
	if(admin == 1)
	{
			WScript.Echo("User is in Local Administrators group\nUse BypassUAC")
	}
}
	
/////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////

// check we can modify the file
function funcWritable(file)
{
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	// couldnt find a function to check write permission so lets just try moving a file and seeing if it fails 
	try
	{
		fso.MoveFile(file,file)
		return 1;
	}
	// if move failed then we caint write there
	catch (e) {return 0;}
}


//Win32_Service get services put into 2d array array[dictionary,dictionary,dictionary]
//////////////////////////////////////////////////////////////////////////
function funcWmiService()
{
	var loc = new ActiveXObject("WbemScripting.SWbemLocator");
	var objWMIService = loc.ConnectServer(".", "root\\cimv2");
	var colItems = objWMIService.ExecQuery("Select * from  Win32_Service");
	var items = new Enumerator(colItems);
	var dictionary = {};
	var arrayDictionary = new Array();
	
	var i=0;
	while (!items.atEnd())
	{
		dictionary = {};
		dictionary.AcceptPause = items.item().AcceptPause;
		dictionary.AcceptStop = items.item().AcceptStop;
		dictionary.Caption = items.item().Caption;
		dictionary.CheckPoint = items.item().CheckPoint;
		dictionary.CreationClassName = items.item().CreationClassName;
		dictionary.Description = items.item().Description;
		dictionary.DesktopInteract = items.item().DesktopInteract;
		dictionary.DisplayName = items.item().DisplayName;
		dictionary.ErrorControl = items.item().ErrorControl;
		dictionary.ExitCode = items.item().ExitCode;
		dictionary.Name = items.item().Name;
		dictionary.PathName = items.item().PathName;
		dictionary.ProcessId = items.item().ProcessId;
		dictionary.ServiceSpecificExitCode = items.item().ServiceSpecificExitCode;
		dictionary.Started = items.item().Started;
		dictionary.StartMode = items.item().StartMode;
		dictionary.StartName = items.item().StartName;
		dictionary.State = items.item().State;
		dictionary.Status = items.item().Status;
		dictionary.SystemCreationClassName = items.item().SystemCreationClassName;
		dictionary.SystemName = items.item().SystemName;
		dictionary.TagId = items.item().TagId;
		dictionary.WaitHint = items.item().WaitHint;
			
		//Multi Array [row][column]
		arrayDictionary[i] = dictionary;
		
		items.moveNext();
		i++;
	}
	return arrayDictionary;
}

// check all services and see if the binary can be overwritten
function funcServiceWritableBinary()
{
		try
		{
			var arrayDictionary = new Array();
			arrayDictionary = funcWmiService();
			
			// loop array of services
			for (dictionary in arrayDictionary)
			{
				//get path to service binary
				var path = arrayDictionary[dictionary]["PathName"];
				
				// split string and get the full file name
				// if quotes in string split string
				path = path.split("\"");
				
				// check if first index in split string is empty if its not then there were quotes in our string
				if(path[0].length > 1)
				{
					// split on space to get rid of arguments
					path = path[0].split(" ");
					path = path[0];	
				}
				// there were quotes in the string so grabe the second index 
				else
					path = path[1];
				
				// check if path is writeable
				var isWriteable = funcWritable(path);
				if(isWriteable)
				{
					WScript.Echo("Name: " + arrayDictionary[dictionary]["DisplayName"]);
					WScript.Echo("Path: " + path);
					WScript.Echo("Start Type: " + arrayDictionary[dictionary]["StartMode"]);
					WScript.Echo('\n');
				}
				
			}
		}
		catch (e){}
		
	
		
}
function funcUnquotedServicePath()
{
	try
	{
		var arrayDictionary = new Array();
		arrayDictionary = funcWmiService();
		
		// loop array of services
		for (dictionary in arrayDictionary)
		{
			//get path to service binary
			var path = arrayDictionary[dictionary]["PathName"];
			
			// split string and get the full file name
			// if no quotes in string 
			if(path.indexOf('"') <= -1)
			{
			
				// split on . to get rid of arguments
				path = path.split(".exe");
				//check for space in path which means its vunerable
				if(path[0].indexOf(' ') > -1)
				{
					WScript.Echo("Name: " +arrayDictionary[dictionary]["DisplayName"]);
					WScript.Echo("Path: " +arrayDictionary[dictionary]["PathName"]);
					WScript.Echo("Start Type: "+arrayDictionary[dictionary]["StartMode"]);
					WScript.Echo('\n');
				}
			}
		}
	}
	catch (e){}
	
}

///////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////


// Enumerate all key value names
////////////////////////////////////////////////////////////////////////
function funcEnumKeyValues(root,path)
{
	// use WMI to enum reg keys
	var loc = new ActiveXObject("WbemScripting.SWbemLocator");
	var objWMIService = loc.ConnectServer(".", "root\\default");
	var objReg = objWMIService.Get("StdRegProv");

	//Prepare EnumValues Method
	var objMethod = objReg.METHODS_.Item("EnumValues");
	var objParamsIn = objMethod.InParameters.SpawnInstance_();
	objParamsIn.hDefKey = root;
	objParamsIn.sSubKeyName = path;

	//Execute
	var objParamsOut = objReg.ExecMethod_(objMethod.Name,objParamsIn);

	
		//Convert output to Array
		var paramasOutArray = objParamsOut.sNames.toArray();
		
		return paramasOutArray;
}

// Get Key Value based on Key Name
//////////////////////////////////////////////////////////////////////////////
function funcRegKeyValue(root,path,keys)
{
	var dictionary = {}
	// use WMI to get reg key value
	var loc = new ActiveXObject("WbemScripting.SWbemLocator");
	var objWMIService = loc.ConnectServer(".", "root\\default");
	var objReg = objWMIService.Get("StdRegProv");
	
	// Loop over each key
	for (var i=0;i < keys.length;++i)
	{
		//Prepare getStringValue Method
		var objMethod = objReg.METHODS_.Item("getStringValue");
		var objParamsIn = objMethod.InParameters.SpawnInstance_();
		objParamsIn.hDefKey = root;
		objParamsIn.sSubKeyName = path;
		objParamsIn.sValueName = keys[i];

		//Execute
		var objParamsOut = objReg.ExecMethod_(objMethod.Name,objParamsIn);
		var stringParamsOut = String(objParamsOut.sValue);
		
		dictionary[keys[i]] = stringParamsOut;
	}
	return dictionary;
}


//HKLM-Software-Microsoft-Windows-CurrentVersion-Run
//////////////////////////////////////////////////////////////////////
function funcHKLM_Software_Microsoft_Windows_CurrentVersion_Run()
{
	dictionary = {};
	try
	{
		var HKLM = 0x80000002; //HKCU = 0x80000001; 
		var path = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
		
		keys = funcEnumKeyValues(HKLM,path);
		dictionary = funcRegKeyValue(HKLM,path,keys);
		
		for(key in dictionary)
		{
			binPath = dictionary[key].replace('"','').replace('"','');
			// check if path is writeable
			var isWriteable = funcWritable(binPath);
			if(isWriteable)
			{
				WScript.Echo(binPath);
			}
			
		}
		
		
	}
	catch (e){}
	
}

//HKLM-Software-Microsoft-Windows-CurrentVersion-RunOnce
//////////////////////////////////////////////////////////////////////
function funcHKLM_Software_Microsoft_Windows_CurrentVersion_RunOnce()
{
	dictionary = {};
	try
	{
		var HKLM = 0x80000002; //HKCU = 0x80000001; 
		var path = "Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce";
		
		keys = funcEnumKeyValues(HKLM,path);
		dictionary = funcRegKeyValue(HKLM,path,keys);
		
		
		for(key in dictionary)
		{
			binPath = dictionary[key].replace('"','').replace('"','');
			// check if path is writeable
			var isWriteable = funcWritable(binPath);
			if(isWriteable)
			{
				WScript.Echo(binPath);
			}	
		}	
	}
	catch (e){}
	
}

//HKLM-Software-Wow6432Node-Microsoft-Windows-CurrentVersion-RunOnce
//////////////////////////////////////////////////////////////////////
function funcHKLM_Software_Wow6432Node_Microsoft_Windows_CurrentVersion_RunOnce()
{
	dictionary = {};
	try
	{
		var HKLM = 0x80000002; //HKCU = 0x80000001; 
		var path = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\RunOnce";
		
		keys = funcEnumKeyValues(HKLM,path);
		dictionary = funcRegKeyValue(HKLM,path,keys);
		
		for(key in dictionary)
		{
			binPath = dictionary[key].replace('"','').replace('"','');
			// check if path is writeable
			var isWriteable = funcWritable(binPath);
			if(isWriteable)
			{
				WScript.Echo(binPath);
			}
	
		}

	}
	catch (e){}
	
}

//HKLM-Software-Wow6432Node-Microsoft-Windows-CurrentVersion-Run
//////////////////////////////////////////////////////////////////////
function funcHKLM_Wow6432Node_Software_Microsoft_Windows_CurrentVersion_Run()
{
	dictionary = {};
	try
	{
		var HKLM = 0x80000002; //HKCU = 0x80000001; 
		var path = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Run";
		
		keys = funcEnumKeyValues(HKLM,path);
		dictionary = funcRegKeyValue(HKLM,path,keys);
		
		for(key in dictionary)
		{
			binPath = dictionary[key].replace('"','').replace('"','');
			// check if path is writeable
			var isWriteable = funcWritable(binPath);
			if(isWriteable)
			{
				WScript.Echo(binPath);
			}
			
		}
		
		
	}
	catch (e){}
	
}

//HKLM-Software-Microsoft-Windows-CurrentVersion-RunService
//////////////////////////////////////////////////////////////////////
function funcHKLM_Software_Microsoft_Windows_CurrentVersion_RunService()
{
	dictionary = {};
	try
	{
		var HKLM = 0x80000002; //HKCU = 0x80000001; 
		var path = "Software\\Microsoft\\Windows\\CurrentVersion\\RunService";
		
		keys = funcEnumKeyValues(HKLM,path);
		dictionary = funcRegKeyValue(HKLM,path,keys);
		
		for(key in dictionary)
		{
			binPath = dictionary[key].replace('"','').replace('"','');
			// check if path is writeable
			var isWriteable = funcWritable(binPath);
			if(isWriteable)
			{
				WScript.Echo(binPath);
			}
			
		}
		
		
	}
	catch (e){}
	
}

//HKLM-Software-Microsoft-Windows-CurrentVersion-RunOnceService
//////////////////////////////////////////////////////////////////////
function funcHKLM_Software_Microsoft_Windows_CurrentVersion_RunOnceService()
{
	dictionary = {};
	try
	{
		var HKLM = 0x80000002; //HKCU = 0x80000001; 
		var path = "Software\\Microsoft\\Windows\\CurrentVersion\\RunOnceService";
		
		keys = funcEnumKeyValues(HKLM,path);
		dictionary = funcRegKeyValue(HKLM,path,keys);
		
		for(key in dictionary)
		{
			binPath = dictionary[key].replace('"','').replace('"','');
			// check if path is writeable
			var isWriteable = funcWritable(binPath);
			if(isWriteable)
			{
				WScript.Echo(binPath);
			}
			
		}
		
		
	}
	catch (e){}
	
}
//HKLM-Software-Wow6432Node-Microsoft-Windows-CurrentVersion-RunService
//////////////////////////////////////////////////////////////////////
function funcHKLM_Wow6432Node_Software_Microsoft_Windows_CurrentVersion_RunService()
{
	dictionary = {};
	try
	{
		var HKLM = 0x80000002; //HKCU = 0x80000001; 
		var path = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\RunService";
		
		keys = funcEnumKeyValues(HKLM,path);
		dictionary = funcRegKeyValue(HKLM,path,keys);
		
		for(key in dictionary)
		{
			binPath = dictionary[key].replace('"','').replace('"','');
			// check if path is writeable
			var isWriteable = funcWritable(binPath);
			if(isWriteable)
			{
				WScript.Echo(binPath);
			}
			
		}
		
		
	}
	catch (e){}
	
}
//HKLM-Software-Wow6432Node-Microsoft-Windows-CurrentVersion-RunOnceService
//////////////////////////////////////////////////////////////////////
function funcHKLM_Wow6432Node_Software_Microsoft_Windows_CurrentVersion_RunOnceService()
{
	dictionary = {};
	try
	{
		var HKLM = 0x80000002; //HKCU = 0x80000001; 
		var path = "Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\RunOnceService";
		
		keys = funcEnumKeyValues(HKLM,path);
		dictionary = funcRegKeyValue(HKLM,path,keys);
		
		for(key in dictionary)
		{
			binPath = dictionary[key].replace('"','').replace('"','');
			// check if path is writeable
			var isWriteable = funcWritable(binPath);
			if(isWriteable)
			{
				WScript.Echo(binPath);
			}
			
		}
		
		
	}
	catch (e){}
	
}

/////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////

function showSubFolders(Path)
{
		
	var FSO = new ActiveXObject("Scripting.FileSystemObject");
	var Folder = FSO.GetFolder(Path);
	var sfs = new Enumerator(Folder.SubFolders);

	for(;!sfs.atEnd(); sfs.moveNext())
	{	
		
		folders_Schedule_Tasks.push(sfs.item());// add folder path to global array
		showSubFolders(sfs.item())// recusivly search all folders and subfolders
	}
	return;
}

function showFiles(Path)
{
	var FSO = new ActiveXObject("Scripting.FileSystemObject");
	var Folder = FSO.GetFolder(Path);
	var sfs = new Enumerator(Folder.Files);

	for(;!sfs.atEnd(); sfs.moveNext())
	{	
		//WScript.Echo(sfs.item());
		files_Schedule_Tasks.push(sfs.item());// add folder path to global array
	}
	return;
}

function getScheduleTasksFiles()
{
	// path where schdules tasks files are held (xml format)
	showSubFolders("C:\\Windows\\System32\\Tasks");
	for (folder in folders_Schedule_Tasks)
	{
		try
		{
			showFiles(String(folders_Schedule_Tasks[folder]));
		}
		catch (e){}
	}
	for (file in files_Schedule_Tasks)
	{
		
		
		try
		{
			
			rline = new Array();
			fs = new ActiveXObject("Scripting.FileSystemObject");
			var oShell = new ActiveXObject("WScript.Shell");
			oShell.Run("cmd.exe /c type \"" +files_Schedule_Tasks[file] + "\" > tmpread.txt;exit", 0, false);
			WScript.Sleep(1000);
			// read file
			f = fs.GetFile("tmpread.txt");
			
			is = f.OpenAsTextStream(1,0);
			var count =0;
			while(!is.AtEndOfStream)
			{
					rline[count]= is.ReadLine();
					count++;
			}
			is.Close();
			
			
			var xmlTxt = "";
			for(i=0;i<rline.length;i++)
			{
					xmlTxt += rline[i] + "\n";
					
			}
			
			//WScript.Echo(xmlTxt);
			// check for command 
			var command = ""
			command = xmlTxt.split("<Command>");
			command = command[1].split("</Command>");
			
			
			var isWriteable = funcWritable(command[0].replace('"','').replace('"',''));
			if(isWriteable)
			{
				WScript.Echo(files_Schedule_Tasks[file]);
				WScript.Echo(command[0].replace('"','').replace('"','') + "\n");
			}
		}
		catch (e){}
		oShell.Run("cmd.exe /c del tmpread.txt", 0, false);	
	}
}

//bypass uac
WScript.Echo("\n***********User is in admin group***********");
funcLocalAdminGroup();

// modify registry run key binary
WScript.Echo("\n***********Auto Run Registry Writable Binary***********");

WScript.Echo("\n            HKLM Run ");
funcHKLM_Software_Microsoft_Windows_CurrentVersion_Run();
WScript.Echo("\n            HKLM Once ");
funcHKLM_Software_Microsoft_Windows_CurrentVersion_RunOnce();
WScript.Echo("\n            HKLM Wow6432Node Run ");
funcHKLM_Wow6432Node_Software_Microsoft_Windows_CurrentVersion_Run();
WScript.Echo("\n            HKLM Wow6432Node RunOnce ");
funcHKLM_Software_Wow6432Node_Microsoft_Windows_CurrentVersion_RunOnce();
WScript.Echo("\n            HKLM RunService ");
funcHKLM_Software_Microsoft_Windows_CurrentVersion_RunService();
WScript.Echo("\n            HKLM RunOnceService ");
funcHKLM_Software_Microsoft_Windows_CurrentVersion_RunOnceService();
WScript.Echo("\n            HKLM Wow6432Node RunService ");
funcHKLM_Wow6432Node_Software_Microsoft_Windows_CurrentVersion_RunService();
WScript.Echo("\n            HKLM Wow6432Node RunOnceSrvice ");
funcHKLM_Wow6432Node_Software_Microsoft_Windows_CurrentVersion_RunOnceService();

// unquoted service path
WScript.Echo("\n***********Unquoted Service Path***********");
funcUnquotedServicePath();
//Service binary can be over written
WScript.Echo("\n***********Writable Service Binary***********");
funcServiceWritableBinary();

// scheduled tasks binary can be over written
WScript.Echo("\n***********Scheduled Tasks Writable Binary***********");
var folders_Schedule_Tasks = new Array("C:\\Windows\\System32\\Tasks");
var files_Schedule_Tasks = new Array("");
getScheduleTasksFiles();
