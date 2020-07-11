/*******************************************************************************
 custom_vspaths.js - change Visual Studio paths for default location of projects,
                     settings etc.

 This file is free for use as public domain by anyone under the terms of
 Creative Commons Attribution-NonCommercial 4.0 International Public License
 See the the file license.txt and https://creativecommons.org

 originally created by Jerker Back 2006
 
  https://github.com/jerker-back/VSUtilities

================================================================================
 RcsID = $Id$ */

// HowTo: Run as Administrator and execute "custom_vspaths.js"

/*
Comment and background:
This script is a help on the way to move Visual Studio out of My Documents.
It was made due to the habit of mine of NOT using the "My Documents"
directory for development. Instead I have a dedicated drive for these things.
Being part of Windows installation and a "Special folder", "My Documents" are
usually filled with all sorts of installation folders. It is by me treated as a
per Windows installation existing folder and I make sure not to store anything
important I want to keep there. 

*/

var WshShell = WScript.CreateObject("WScript.Shell");
var objEnv = WshShell.Environment("Volatile");
var g_strBaseKey = "HKCU\\Software\\Microsoft\\";

var g_user = objEnv("USERNAME");
var g_strMyDocumentsLocation = WshShell.SpecialFolders("MyDocuments");

// NOTE; Change these as appropriate
var g_strProjectLocation = "F:\\dev\\Projekt";              // persistent place for project files, source code etc
var g_strVisualStudioLocation = "D:\\devapps\\VisualStudio";  // persistent place for user macros, settings, snippets etc

// For me, this will give the following locations for user created things, 
// also common to and shared by all Visual Studio versions installed.
// F:\\dev\\Projekt
// F:\Users\VisualStudio\Settings\vs2005.vssettings
// F:\Users\VisualStudio\Addins
// F:\Users\VisualStudio\Code Snippets
// ... etc
// F:\Users\VisualStudio\VSMacros\8.0\samples.vsmacros
// F:\Users\VisualStudio\VSMacros\<user>.vsmacros

// The settinsg files are renamed to vsXXXX..vssettings, where XXXX is the
// Visual Studio year version. Express versions can be added to the script
// if needed

main();
WScript.Echo("Visual Studio user paths successfully changed!");
WScript.Echo("Now, move your settings files to the new location\nExample: current.vssettings => vs2005.vssettings");

// Visual Studio 2003 - 2013
// SQL Server Management Studio 2008 - 2014

function main()
{
	// Find Visual Studio 7.1, 8.0, 9.0, 10.0, 11.0, 12.0
	var bVS71 = null;
	var bVS80 = null;
	var bVS90 = null;
	var bVS100 = null;
	var bVS110 = null;
	var bVS120 = null;
	var bVS140 = null;
	var bVS80X = null;
	var bVS90X = null;
	var bVS100X = null;
	var bVS110X = null;
	var bVS120X = null;
	var bMSDN80 = null;
	var bMSDN90 = null;
	var bVSA80 = null;
	var bVSA90 = null;
	var bSQL100 = null;
	var bSQL110 = null;
	var bSQL120 = null;

	try { bVS71 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\7.1\\MyDocumentsLocation"); }   catch (e) { bVS71 = null;   }
	try { bVS80 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\8.0\\MyDocumentsLocation"); }   catch (e) { bVS80 = null;   }
	try { bVS90 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\9.0\\MyDocumentsLocation"); }   catch (e) { bVS90 = null;   }
	try { bVS100 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\10.0\\MyDocumentsLocation"); } catch (e) { bVS100 = null;  }
	try { bVS110 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\11.0\\MyDocumentsLocation"); } catch (e) { bVS110 = null;  }
	try { bVS120 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\12.0\\MyDocumentsLocation"); } catch (e) { bVS120 = null;  }
	try { bVS140 = WshShell.RegRead(g_strBaseKey + "VisualStudio\\14.0\\MyDocumentsLocation"); } catch (e) { bVS140 = null; }
	try { bVS80X = WshShell.RegRead(g_strBaseKey + "VCExpress\\8.0\\MyDocumentsLocation"); } 	 catch (e) { bVS80X = null; }
	try { bVS90X = WshShell.RegRead(g_strBaseKey + "VCExpress\\9.0\\MyDocumentsLocation"); }     catch (e) { bVS90X = null;  }
	try { bVS100X = WshShell.RegRead(g_strBaseKey + "VCExpress\\10.0\\MyDocumentsLocation"); }   catch (e) { bVS100X = null; }
	try { bVS110X = WshShell.RegRead(g_strBaseKey + "VCExpress\\11.0\\MyDocumentsLocation"); }   catch (e) { bVS110X = null; }
	try { bVS120X = WshShell.RegRead(g_strBaseKey + "VCExpress\\12.0\\MyDocumentsLocation"); }   catch (e) { bVS120X = null; }
	try { bMSDN80 = WshShell.RegRead(g_strBaseKey + "MSDN\\8.0\\MyDocumentsLocation"); }         catch (e) { bMSDN80 = null; }
	try { bMSDN90 = WshShell.RegRead(g_strBaseKey + "MSDN\\9.0\\MyDocumentsLocation"); }         catch (e) { bMSDN90 = null; }
	try { bVSA80 = WshShell.RegRead(g_strBaseKey + "VSA\\8.0\\MyDocumentsLocation"); }           catch (e) { bVSA80 = null;  }
	try { bVSA90 = WshShell.RegRead(g_strBaseKey + "VSA\\9.0\\MyDocumentsLocation"); }           catch (e) { bVSA90 = null;  }
	try { bSQL100 = WshShell.RegRead(g_strBaseKey + "Microsoft SQL Server\\100\\Tools\\Shell\\MyDocumentsLocation"); } catch (e) { bSQL100 = null; }
	try { bSQL110 = WshShell.RegRead(g_strBaseKey + "SQL Server Management Studio\\11.0\\MyDocumentsLocation"); } catch (e) { bSQL110 = null; }
	try { bSQL120 = WshShell.RegRead(g_strBaseKey + "SQL Server Management Studio\\12.0\\MyDocumentsLocation"); } catch (e) { bSQL120 = null; }

	try
	{
		if (bVS71) RegisterCustomVSPaths("VisualStudio\\",   "7.1", "vs2003", true);
		if (bVS80) RegisterCustomVSPaths("VisualStudio\\",   "8.0", "vs2005", true);
		if (bVS90) RegisterCustomVSPaths("VisualStudio\\",   "9.0", "vs2008", true);
		if (bVS100) RegisterCustomVSPaths("VisualStudio\\", "10.0", "vs2010", true);
		if (bVS110) RegisterCustomVSPaths("VisualStudio\\", "11.0", "vs2012", true);
		if (bVS120) RegisterCustomVSPaths("VisualStudio\\", "12.0", "vs2013", true);
		if (bVS140) RegisterCustomVSPaths("VisualStudio\\", "14.0", "vs2015", true);
		if (bVS80X) RegisterCustomVSPaths("VCExpress\\", "8.0", "vcx2005", false);
		if (bVS90X) RegisterCustomVSPaths("VCExpress\\",   "9.0", "vcx2008", false);
		if (bVS100X) RegisterCustomVSPaths("VCExpress\\", "10.0", "vcx2010", false);
		if (bVS110X) RegisterCustomVSPaths("VCExpress\\", "11.0", "vcx2012", false);
		if (bVS110X) RegisterCustomVSPaths("VCExpress\\", "12.0", "vcx2013", false);
		if (bMSDN80) RegisterCustomVSPaths("MSDN\\", "8.0", null, false );
		if (bMSDN90) RegisterCustomVSPaths("MSDN\\", "9.0", null, false);
		if (bVSA80) RegisterCustomVSPaths("VSA\\", "8.0", null, false);
		if (bVSA90) RegisterCustomVSPaths("VSA\\", "9.0", null, false);
		if (bSQL100) RegisterCustomVSPaths("Microsoft SQL Server\\100\\Tools\\Shell\\","", "mssql2008", false);
		if (bSQL110) RegisterCustomVSPaths("SQL Server Management Studio\\11.0\\","", "mssql2012", false);
		if (bSQL120) RegisterCustomVSPaths("SQL Server Management Studio\\12.0\\","", "mssql2014", false);
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
	    // Projects
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

		// Settings file "vssettings"
		strKey = "\\Profile\\";

		if (strVSName != null && strVSName != "")
			WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "AutoSaveFile", "%vsspv_visualstudio_dir%\\Settings\\" + strVSName + ".vssettings", "REG_EXPAND_SZ");


        // Visual Studio macros file
		strKey = "\\vsmacros\\";
		if (bRegMacros) 
		{
			WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "OtherProjects7\\0\\Path", g_strVisualStudioLocation + "\\VSMacros\\" + strVSVer + "\\samples.vsmacros", "REG_SZ");
			WshShell.RegWrite(g_strBaseKey + strVSProduct + strVSVer + strKey + "RecordingProject7\\Path", g_strVisualStudioLocation + "\\VSMacros\\" + g_user + ".vsmacros", "REG_SZ");
		}
	}
	catch (e)
	{
		WScript.Echo("ERROR: Cannot write to registry.");
		return null;
	}

}

