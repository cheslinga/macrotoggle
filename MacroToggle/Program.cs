using System;
using Microsoft.Win32;
using System.Windows.Forms;

namespace MacroToggle
{
    class Program
    {
        const Int32 ENABLE = 0x1;
        const Int32 DISABLE = 0x4;

        private static void Main(string[] args)
        {
            string msg;
            KeyReference regkey = new KeyReference();

            //Check that the subkey name is not null.
            if (regkey.SubKeyName != null)
            {
                //Check that the subkey's value is not null on top of that.
                if (regkey.dword_val != null)
                {   
                    //Enable if disabled, disable if enabled.
                    if ((Int32)regkey.dword_val == ENABLE)
                    {
                        regkey.SetValue(DISABLE);
                        msg = "Macros have been disabled!";
                    } else
                    {
                        regkey.SetValue(ENABLE);
                        msg = "Macros have been enabled!";
                    }
                } else
                {
                    msg = "ERROR: the registry subkey " + regkey.SubKeyName + "\\" + regkey.ValName + " could not be found...";
                }
            } else
            {
                msg = "ERROR: the registry key for Microsoft Excel could not be found.";
            }

            //Print the msg variable
            DialogResult _ = MessageBox.Show(msg, "MacroToggle", MessageBoxButtons.OK);
        }
    }
    
    //Wrapper class to read/write to the Windows registry.
    public class KeyReference
    {
        const string VAL_NAME = "VBAWarnings";
        static readonly string[] EXCEL_SUBKEYS =
        {
            "SOFTWARE\\Microsoft\\Office\\16.0\\Excel\\Security",
            "SOFTWARE\\Microsoft\\Office\\14.0\\Excel\\Security",
            "SOFTWARE\\Microsoft\\Office\\12.0\\Excel\\Security"
        };

        public string SubKeyName { get; set; }
        public string ValName { get { return VAL_NAME; } }

        static RegistryHive hive;
        static RegistryKey macros_key;
        public readonly Object dword_val;

        public KeyReference()
        {
            hive = RegistryHive.CurrentUser;
            SubKeyName = SearchForSubkeys();

            if (SubKeyName != null)
            {
                macros_key = RegistryKey.OpenBaseKey(hive, RegistryView.Default).OpenSubKey(SubKeyName, true);
                dword_val = macros_key.GetValue(VAL_NAME);
            }
        }

        //Sets the value of the VBAWarnings registry value.
        public void SetValue(Int32 value)
        {
            macros_key.SetValue(VAL_NAME, value);
        }
        
        //Checks the HKCU hive for each key in the subkeys list.
        private string SearchForSubkeys()
        {
            foreach (string key in EXCEL_SUBKEYS)
            {   
                //Check that attempting to open the key does not return null.
                if (RegistryKey.OpenBaseKey(hive, RegistryView.Default).OpenSubKey(key, false) != null)
                {
                    return key;
                }
            }
            //If not key was returned, return a null.
            return null;
        }
    }
}
