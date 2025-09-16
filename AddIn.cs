using Extensibility;
using Microsoft.Office.Core;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OnenoteAddin
{
    [ComVisible(true)]
    [Guid("90F757AF-A542-4573-884E-A72CB573F04E")] // このGUIDをregasmとinstall.regで使用
    [ProgId("OnenoteAddin.AddIn")]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private Microsoft.Office.Interop.OneNote.Application oneNoteApp;
        private object addInInstance;
        private OneNoteOperator oneNoteOperator;

        #region IDTExtensibility2 インターフェイスの実装
        public void OnConnection(object Application,
                                 ext_ConnectMode ConnectMode,
                                 object AddInInst,
                                 ref Array custom)
        {
            this.oneNoteApp = (Microsoft.Office.Interop.OneNote.Application)Application;
            this.addInInstance = AddInInst;
            this.oneNoteOperator = new OneNoteOperator(this.oneNoteApp);
        }

        public void OnStartupComplete(ref Array custom) { }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            this.oneNoteApp = null;
            this.oneNoteOperator = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion

        #region IRibbonExtensibility インターフェイスの実装
        public string GetCustomUI(string RibbonID)
        {
            // ここでリボンのUIを定義するXML文字列を返します。
            return @"<?xml version=""1.0"" encoding=""utf-8""?>
                    <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
                        <ribbon>
                            <tabs>
                                <tab idMso='TabHome'>
                                    <group id='GroupMD' label='MD'>
                                        <button id='ConvertButton'
                                                size='normal'
                                                showLabel='false'
                                                screentip='MD:選択範囲を変換'
                                                imageMso='AcetateModeOriginalMarkup'
                                                onAction='OnConvertButtonClick'/>
                                    </group>
                                </tab>
                            </tabs>
                        </ribbon>
                    </customUI>";
        }
        #endregion

        #region 内部処理の実装
        public void OnConvertButtonClick(IRibbonControl control)
        {
            try
            {
                string selectedText = oneNoteOperator.GetSelectedText();
                if (string.IsNullOrEmpty(selectedText))
                {
                    return;
                }

                var stylefile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "OnenoteAddin", "style.css");
                string style = File.ReadAllText(stylefile);
                string body = MarkdownOperator.ConvertMarkdownToHtml(selectedText);
                oneNoteOperator.ReplaceSelectedTextWithHtmlBlock(style, body);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("エラーが発生しました: \n" + ex.Message);
                AddIn.WriteErrorLog(ex.ToString());
            }
        }

        public static void WriteDebugLog(string message)
        {
            var logfile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "OnenoteAddin", "debug.log");
            File.AppendAllText(logfile, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message + "\n");
        }

        public static void WriteErrorLog(string message)
        {
            var logfile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "OnenoteAddin", "error.log");
            File.AppendAllText(logfile, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message + "\n");
        }
        #endregion
    }
}
