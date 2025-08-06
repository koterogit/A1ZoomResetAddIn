using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace A1ZoomResetAddIn
{
    public partial class ThisAddIn
    {
        internal void ResetA1AndZoomAllSheets()
        {
            Excel.Application app = this.Application;
            Excel.Workbook wb = app?.ActiveWorkbook;
            if (app == null || wb == null) return;

            bool prevScreenUpdating = app.ScreenUpdating;
            try
            {
                app.ScreenUpdating = false;

                var originalSheet = app.ActiveSheet as Excel.Worksheet;

                foreach (var sheetObj in wb.Worksheets)
                {
                    if (sheetObj is Excel.Worksheet ws)
                    {
                        try
                        {
                            ws.Activate();
                            Excel.Range a1 = ws.Range["A1"];
                            a1.Select();
                            if (app.ActiveWindow != null)
                            {
                                app.ActiveWindow.Zoom = 100;
                            }
                        }
                        catch { /* ワークシート個別の例外は握りつぶす */ }
                        finally
                        {
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws);
                        }
                    }
                    else
                    {
                        // 図表シート等はスキップ
                    }
                }

                originalSheet?.Activate();
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
            }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // 起動時に自動実行したい場合は下記をアンコメント
            // ResetA1AndZoomAllSheets();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
