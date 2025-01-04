using System;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using EdmLib;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace EPDMCopyReferencesToClipboard
{
    [Guid("6E4F7AE3-6210-45D8-A964-B7D14A65C80E"), ComVisible(true)]
    public class CopyRefsToClipboardAddin : IEdmAddIn5
    {
       
        /// <summary>
        /// SolidWorks Enterprise PDM calls this method when an add-in is loaded to retrieve information
        /// about the add-in and the commands it supports. You supply the information by implementing the routine.
        /// </summary>
        /// <param name="poInfo">Return information about your add-in in the members of this struct. Most 
        /// of the data will be displayed in the Administrate Add-ins dialog box.</param>
        /// <param name="vault">Pointer to the active file vault.</param>
        /// <param name="cmdMgr">The command manager. You use this pointer to add hooks, menu commands and toolbar buttons.</param>
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 vault, IEdmCmdMgr5 cmdMgr)
        {

            poInfo.mbsAddInName = "PDM Copy References To Clipboard.";
            poInfo.mbsCompany = "EhCanadian Consulting, Inc.";
            poInfo.mbsDescription = "Copies filenames of references of selected files to the clipboard, space delimited.";
            poInfo.mlAddInVersion = 1;
            poInfo.mlRequiredVersionMajor = 7;
            poInfo.mlRequiredVersionMinor = 0;

            int cmdID = 31516252; // C=3, O=15, P=16, Y=25
            string menuString = "Copy references to clipboard";
            int edmMenuFlags = 3;
            string statusbarHelp = "Copy references to clipboard.";
            string toolbarTooltip = "Copy references to clipboard.";
            int toolbarButtonIndex = 0;
            int toolbarImageId = 0;

            cmdMgr.AddCmd(cmdID, menuString, edmMenuFlags, statusbarHelp, toolbarTooltip, toolbarButtonIndex, toolbarImageId);

            cmdID = 31516253; // C=3, O=15, P=16, Y=25
            menuString = "Copy where used to clipboard";
            statusbarHelp = "Copy where used to clipboard.";
            toolbarTooltip = "Copy where used to clipboard.";

            cmdMgr.AddCmd(cmdID, menuString, edmMenuFlags, statusbarHelp, toolbarTooltip, toolbarButtonIndex, toolbarImageId);

            cmdID = 31516254; // C=3, O=15, P=16, Y=25
            menuString = "Copy upper level drawings to clipboard";
            statusbarHelp = "Copy upper level drawings to clipboard.";
            toolbarTooltip = "Copy upper level drawings to clipboard.";

            cmdMgr.AddCmd(cmdID, menuString, edmMenuFlags, statusbarHelp, toolbarTooltip, toolbarButtonIndex, toolbarImageId);

        }

        /// <summary>
        /// SolidWorks Enterprise PDM will call this method whenever one of the menu
        /// commands or hooks that are registered in IEdmAddIn5::GetAddInInfo is executed.
        /// </summary>
        /// <param name="edmCmd">Command information common to all affected files and folders.</param>
        /// <param name="data">An array with one struct per affected file or folder.</param>
        public void OnCmd(ref EdmCmd edmCmd, ref EdmCmdData[] data)
        {

            if (data == null)
                return;

            EdmVault5 vault = (EdmVault5)edmCmd.mpoVault;
            List<string> referencedFiles = new List<string>();
            try
            {
                for (int i = 0; i < data.Length; i++)
                {
                    EdmCmdData cmdData = (EdmCmdData)data.GetValue(i);
                    int fileId = cmdData.mlObjectID1;
                    int parentFolderId = cmdData.mlObjectID3;
                    string filename = cmdData.mbsStrData1;

                    IEdmFile5 fileToProcess = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, fileId);
                    if (fileToProcess != null)
                    {
                        string projectName = "";
                        IEdmReference5 reference = fileToProcess.GetReferenceTree(parentFolderId, 0);
                        switch (edmCmd.mlCmdID)
                        {
                            case 31516252:
                                GetReferences(ref referencedFiles, reference, true, ref projectName);
                                break;
                            case 31516253:
                                GetWhereUsed(ref referencedFiles, reference, true);
                                break;
                            case 31516254:
                                GetWhereUsed(ref referencedFiles, reference, false);
                                break;
                            default:
                                break;
                        }
                        reference = null;
                    }
                }
                if (referencedFiles.Count > 0)
                {
                    try
                    {
                        //Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(ThreadException);
                        //AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
                        Kolibri.Clippy.PushStringToClipboard(string.Join(" ", referencedFiles.ToArray()));
                        //vault.MsgBox(edmCmd.mlParentWnd, "Found " + referencedFiles.Count + " file(s).", EdmMBoxType.EdmMbt_Icon_Information, "Copied to clipboard");
                        //MessageBox.Show("Found " + referencedFiles.Count + " file(s).", "Copied to clipboard", MessageBoxButtons.OK);
                        // referencedFiles = null;
                    }
                    finally
                    {
                        //Application.ThreadException -= ThreadException;
                        //AppDomain.CurrentDomain.UnhandledException -= CurrentDomain_UnhandledException;
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string GetReferences(ref List<string> referencedFiles, IEdmReference5 reference, bool topLevel, ref string projectName)
        {
            IEdmPos5 position = reference.GetFirstChildPosition(ref projectName, topLevel, true, 0);
            while (!position.IsNull)
            {
                reference = reference.GetNextChild(position);
                referencedFiles.Add(reference.Name);
                GetReferences(ref referencedFiles, reference, false, ref projectName);
            }
            return projectName;
        }

        private void GetWhereUsed(ref List<string> referencedFiles, IEdmReference5 reference, bool recursive)
        {
            IEdmPos5 position = reference.GetFirstParentPosition(0, false);
            while (!position.IsNull)
            {
                reference = reference.GetNextParent(position);
                // if we only want top level, only add drawings
                if (reference.Name.ToLower().Contains(".slddrw") && !recursive)
                {
                    referencedFiles.Add(reference.Name);
                }
                else if (recursive)
                {
                    referencedFiles.Add(reference.Name); // add anything, it's recursive
                    GetWhereUsed(ref referencedFiles, reference, recursive);
                }
            }
        }

        void ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            var x = e.Exception;
        }
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                var ex = e.ExceptionObject as Exception;
                if (ex == null)
                {
                    ex = new NotSupportedException("Unhandled exception doesn't derive from System.Exception: "
         + e.ExceptionObject.ToString()
      );
                }               
            }
            catch (Exception exc)
            {
               // log.Fatal("Fatal Non-UI Error", exc);
                //try
                //{
                //    MessageBox.Show("Fatal Non-UI Error",
                //        "Fatal Non-UI Error. Could not write the error to the event log. Reason: "
                //        + exc.Message, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //}
                //finally
                //{
                //    Application.Exit();
                //}
            }
        }
        private void CopyToClipboard(string textToCopy)
        {
            Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(ThreadException);
            try
            {
                //Clipboard.Clear();
                //System.Windows.Forms.Clipboard.SetText(textToCopy, System.Windows.Forms.TextDataFormat.Text);
                
                Clipboard.SetDataObject(
        textToCopy, // Text to store in clipboard
        false,       // Do not keep after our application exits
        2,           // Retry 5 times
        200);        // 200 ms delay between retries
            }

            catch (Exception ex)
            {
                var x = 1;
            }
            finally
            {
                //Application.ThreadException -= ThreadException;
            }
        }
    }
}
