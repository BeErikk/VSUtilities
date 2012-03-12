/*******************************************************************************

 Copyright 2006-2010 Jerker Back 

 This work is free under the terms of a 2 clause OSI approved BSD license
 Redistribution and use in source and binary forms, with or without 
 modification, are permitted provided that the following conditions are met:
 
 1 Redistributions of source code must retain the above copyright notice, 
   this list of conditions and the following disclaimer. 
 2 Redistributions in binary form must reproduce the above copyright notice, 
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution. 
 
 THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, 
 THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR 
 PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR 
 CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, 
 EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, 
 PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; 
 OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, 
 WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR 
 OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

================================================================================
  RcsID = $Id$ */

//
// custom_vspaths.js - Customize Visual Studio default paths
// 
//
// Run as Administrator and execute "custom_vspaths.js"

var WshShell = WScript.CreateObject("WScript.Shell");
var objEnv = WshShell.Environment("Volatile");
var g_strBaseKey = "HKCU\\Software\\Microsoft\\";

// NOTE; Change these as appropriate
var g_user = objEnv("USERNAME");
var g_home = objEnv("HOMEDRIVE") + objEnv("HOMEPATH");  // not used
var g_strMyDocumentsLocation = WshShell.SpecialFolders("MyDocuments");
var g_strProjectLocation = "F:\\dev\\Projekt";
var g_strVisualStudioLocation = "F:\\Users\\VisualStudio";
var g_strVSMacrosLocation = g_strVisualStudioLocation + "\\VSMacros";

// Note: This will give the following location of VSMacros
// F:\Users\VisualStudio\VSMacros\8.0\samples.vsmacros
// F:\Users\VisualStudio\VSMacros\jerker.vsmacros

main();
WScript.Echo("Visual Studio user paths successfully changed!");
WScript.Echo("Now, copy your settings files to the new location");

function main()
{
 	// Find Visual Studio 7.1, 8.0, 9.0 and 10.0
	var bVS71 = null;
	var bVS80 = null;
	var bVS90 = null;
	var bVS100 = null;
	var bVS110 = null;
	var bVS80X = null;
	var bVS90X = null;
	var bVS100X = null;
	var bVS110X = null;
	var bMSDN80 = null;
	var bMSDN90 = null;
	var bVSA80 = null;
	var bVSA90 = null;
	var bSQL100 = null;

	try { bVS71 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\7.1\\MyDocumentsLocation"); }   catch (e) { bVS71 = null;   }
	try { bVS80 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\8.0\\MyDocumentsLocation"); }   catch (e) { bVS80 = null;   }
	try { bVS90 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\9.0\\MyDocumentsLocation"); }   catch (e) { bVS90 = null;   }
	try { bVS100 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\10.0\\MyDocumentsLocation"); } catch (e) { bVS100 = null;  }
	try { bVS110 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\11.0\\MyDocumentsLocation"); } catch (e) { bVS110 = null;  }
	try { bVS80X = WshShell.RegRead(g_strBaseKey + "VCExpress\\8.0\\MyDocumentsLocation"); }     catch (e) { bVS80X = null;  }
	try { bVS90X = WshShell.RegRead(g_strBaseKey + "VCExpress\\9.0\\MyDocumentsLocation"); }     catch (e) { bVS90X = null;  }
	try { bVS100X = WshShell.RegRead(g_strBaseKey + "VCExpress\\10.0\\MyDocumentsLocation"); }   catch (e) { bVS100X = null; }
	try { bVS110X = WshShell.RegRead(g_strBaseKey + "VCExpress\\11.0\\MyDocumentsLocation"); }   catch (e) { bVS110X = null; }
	try { bMSDN80 = WshShell.RegRead(g_strBaseKey + "MSDN\\8.0\\MyDocumentsLocation"); }         catch (e) { bMSDN80 = null; }
	try { bMSDN90 = WshShell.RegRead(g_strBaseKey + "MSDN\\9.0\\MyDocumentsLocation"); }         catch (e) { bMSDN90 = null; }
	try { bVSA80 = WshShell.RegRead(g_strBaseKey + "VSA\\8.0\\MyDocumentsLocation"); }           catch (e) { bVSA80 = null;  }
	try { bVSA90 = WshShell.RegRead(g_strBaseKey + "VSA\\9.0\\MyDocumentsLocation"); }           catch (e) { bVSA90 = null;  }
	try { bSQL100 = WshShell.RegRead(g_strBaseKey + "Microsoft SQL Server\\100\\Tools\\Shell\\MyDocumentsLocation"); } catch (e) { bSQL100 = null; }

	try
	{
		if (bVS71) RegisterCustomVSPaths("VisualStudio\\", "7.1", "vs2003", true);
		if (bVS80) RegisterCustomVSPaths("VisualStudio\\", "8.0", "vs2005", true);
		if (bVS90) RegisterCustomVSPaths("VisualStudio\\", "9.0", "vs2008", true);
		if (bVS100) RegisterCustomVSPaths("VisualStudio\\", "10.0", "vs2010", true);
		if (bVS110) RegisterCustomVSPaths("VisualStudio\\", "11.0", "vs2011", true );
		if (bVS80X) RegisterCustomVSPaths("VCExpress\\", "8.0", "vcx2005", false );
		if (bVS90X) RegisterCustomVSPaths("VCExpress\\", "9.0", "vcx2008", false);
		if (bVS100X) RegisterCustomVSPaths("VCExpress\\", "10.0", "vcx2010", false);
		if (bVS110X) RegisterCustomVSPaths("VCExpress\\", "11.0", "vcx2011", false );
		if (bMSDN80) RegisterCustomVSPaths("MSDN\\", "8.0", null, false );
		if (bMSDN90) RegisterCustomVSPaths("MSDN\\", "9.0", null, false);
		if (bVSA80) RegisterCustomVSPaths("VSA\\", "8.0", null, false);
		if (bVSA90) RegisterCustomVSPaths("VSA\\", "9.0", null, false);
		if (bSQL100) RegisterCustomVSPaths("Microsoft SQL Server\\100\\Tools\\Shell\\","", "mssql2008", false);
	}
	catch (e)
	{
		WScript.Echo("ERROR: Cannot write to registry.");
		return null;
	}

	return true;
}

function RegisterCustomVSPaths(strVSProduct, strVSVer, strVSName, bRegMacros)
{
	// Move Visual Studio paths out of My documents location
	var strKey = "\\";
	try
	{
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "DefaultNewProjectLocation", g_strProjectLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "DefaultOpenProjectLocation", g_strProjectLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "DefaultOpenSolutionLocation", g_strProjectLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "DefaultNewProjItemLocation", g_strProjectLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "DefaultOpenProjItemLocation", g_strProjectLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "DefaultFileOpenLocation", g_strProjectLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "DefaultBrowseComponentLocation", g_strProjectLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "VisualStudioProjectsLocation", g_strProjectLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "MyDocumentsLocation", g_strMyDocumentsLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "VisualStudioLocation", g_strVisualStudioLocation, "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "UserItemTemplatesLocation", g_strVisualStudioLocation + "\\Templates\\ItemTemplates", "REG_EXPAND_SZ");
		WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "UserProjectTemplatesLocation", g_strVisualStudioLocation + "\\Templates\\ProjectTemplates", "REG_EXPAND_SZ");
 
		strKey = "\\Profile\\";

		if (strVSName != null && strVSName != "")
			WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "AutoSaveFile", "%vsspv_visualstudio_dir%\\Settings\\" + strVSName + ".vssettings", "REG_EXPAND_SZ");

		strKey = "\\vsmacros\\";

		if (bRegMacros) 
		{
			WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "OtherProjects7\\0\\Path", g_strVSMacrosLocation + "\\" + strVSVer + "\\samples.vsmacros", "REG_SZ");
			WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "RecordingProject7\\Path", g_strVSMacrosLocation + "\\" + g_user + ".vsmacros", "REG_SZ");
		}
	}
	catch (e)
	{
		WScript.Echo("ERROR: Cannot write to registry.");
		return null;
	}

}

