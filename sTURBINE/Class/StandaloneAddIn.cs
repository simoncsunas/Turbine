using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace sTURBINE.Class
{
    class StandaloneAddIn
    {
        static void Main(string[] args)
        {

            SldWorks.SldWorks swApp;

            swApp = new SldWorks.SldWorks();

            swApp.ExitApp();
            swApp = null;
        }

        //Version-specific: HKEY_LOCAL_MACHINE\SOFTWARE\SOLIDWORKS\SOLIDWORKS version\Addins\{CLSID}\Icon Path
        //Version-independent: HKEY_LOCAL_MACHINE\SOFTWARE\SOLIDWORKS\AddIns\{CLSID}\Icon Path
#region SOLIDWORKS Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;

            Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

            string keyname = "SOFTWARE\\SOLIDWORKS\\Addins\\{" + t.GUID.ToString() + "}";

            Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);

            addinkey.SetValue(null, 0);

            addinkey.SetValue("Description", SWattr.Description);

            addinkey.SetValue("Title", SWattr.Title);

#region Extract icon during registration

            BitmapHandler iBmp = new BitmapHandler();

            Assembly thisAssembly;

            thisAssembly = System.Reflection.Assembly.GetExecutingAssembly();

            String tempPath =
            iBmp.CreateFileFromResourceBitmap("_2012_PMP_Interfaces.AddInMgrIcon.bmp",
            thisAssembly);

            // Copy the bitmap to a suitable permanent location with a meaningful filename

            String addInPath = System.IO.Path.GetDirectoryName(thisAssembly.Location);

            String iconPath = System.IO.Path.Combine(addInPath, "AddInMgrIcon.bmp");

            System.IO.File.Copy(tempPath, iconPath, true);

            // Register the icon location

            addinkey.SetValue("Icon Path", iconPath);

#endregion

            keyname = "Software\\SOLIDWORKS\\AddInsStartup\\{" + t.GUID.ToString() + "}";

            addinkey = hkcu.CreateSubKey(keyname);

            addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup),
            Microsoft.Win32.RegistryValueKind.DWord);

        }
#endregion

        //+ API Functionality Dependent on SOLIDWORKS Being Visible
        //Some API functionality is dependent on SOLIDWORKS being visible. For example, if SOLIDWORKS is not visible and you update a drawing and then attempt to update the drawing views, your drawing views might not be updated as expected.
        //To test if SOLIDWORKS must be visible for a specific functionality to work, set SOLIDWORKS to visible and then retry the operation that did not work.
        //-

        //+CommandManager and CommandGroups
        //HKEY_CURRENT_USER\Software\SOLIDWORKS\SOLIDWORKS <version>\User Interface\Custom API Toolbars\<index>
        //HKEY_CURRENT_USER\Software\SOLIDWORKS\SOLIDWORKS <version>\User Interface\Toolbars
        //HKEY_CURRENT_USER\Software\SOLIDWORKS\SOLIDWORKS <version>\User Interface\Toolbars\PartTool
        //HKEY_CURRENT_USER\Software\SOLIDWORKS\SOLIDWORKS <version>\User Interface\Toolbars\AssemblyTool
        //HKEY_CURRENT_USER\Software\SOLIDWORKS\SOLIDWORKS <version>\User Interface\Toolbars\DrawingTool
        //-
    }
}
