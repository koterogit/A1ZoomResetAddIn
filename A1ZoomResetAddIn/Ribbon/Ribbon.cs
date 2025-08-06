using System;
using System.IO;
using System.Reflection;
using Office = Microsoft.Office.Core;

namespace A1ZoomResetAddIn.Ribbon
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
            => GetResourceText("A1ZoomResetAddIn.Ribbon.Ribbon.xml");

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI) => ribbon = ribbonUI;

        public void OnResetClicked(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ResetA1AndZoomAllSheets();
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }
    }
}