using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace WordOverlayProofreader.Addin
{
    [ComVisible(true)]
    public class Ribbon
    {
        private object ribbon;

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordOverlayProofreader.Addin.Ribbon.xml");
        }

        public void Ribbon_Load(object ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void btnScan_Click(object control)
        {
            ThisAddIn.Instance.ScanDocument();
        }
        
        public void btnAutoScan_Click(object control, bool pressed)
        {
            ThisAddIn.Instance.SetAutoScan(pressed);
        }

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
