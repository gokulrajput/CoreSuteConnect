using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using System.Configuration;

namespace CoreSuteConnect.Class.EXIM
{
    class clsDocMaster
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;
        private SAPbouiCOM.ComboBox oComboBox;

        OpenFileDialog OpenFileDialog = new OpenFileDialog();
        string BrowseFilePath = string.Empty;

        #endregion VariableDeclaration

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type = "")
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);
                if (oForm.Mode == BoFormMode.fm_ADD_MODE || Type == "ADDNEWFORM")
                {
                    Form_Load_Components(oForm,"ADD");
                }
            }
            catch (Exception ex)
            {


            }
            finally
            {

            } 
            return BubbleEvent; 
        }


        public bool ItemEvent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            BubbleEvent = true;
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            try
            {
                switch (pVal.EventType)
                { // KEY Down event for multiplication
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {

                        }
                        break;
                    case BoEventTypes.et_FORM_CLOSE:
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {
                            oForm = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        break;
                    case BoEventTypes.et_CLICK:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1")
                            {
                                if (!string.IsNullOrEmpty(Path.GetFileName(BrowseFilePath)))
                                {
                                    oForm.Items.Item("tattach").Specific.value = SBOMain.Get_Attach_Folder_Path() + Path.GetFileName(BrowseFilePath);
                                } 
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1")
                            {
                                //string ReplaceFilePath = "E:\\SAP\\SAP_ATTACHMENTS\\" + FileName;

                                string FileName = Path.GetFileName(oForm.Items.Item("tattach").Specific.value);
                                string ReplaceFilePath = SBOMain.Get_Attach_Folder_Path() + FileName;
                                File.Move(BrowseFilePath, ReplaceFilePath);

                                /*using (Stream s = File.Open(ReplaceFilePath, FileMode.Create))
                                {
                                    using (StreamWriter sw = new StreamWriter(s))
                                    {
                                        sw.Write(BrowseFilePath);
                                    }
                                }*/
                            }
                        }
                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                if (string.IsNullOrEmpty(oForm.Items.Item("tdoccode").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Document Code", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tdoccode").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("tdocname").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Document Name", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tdocname").Click();
                                }
                                else
                                {
                                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cdoctype").Specific;
                                    if (oComboBox.Selected == null)
                                    {
                                        BubbleEvent = false;
                                        SBOMain.SBO_Application.StatusBar.SetText("Please Select Document Type", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        oForm.Items.Item("cdoctype").Click();
                                    }
                                }
                            }
                        }
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                Form_Load_Components(oForm, "ADD");
                            }
                            else if (pVal.ItemUID == "btnATC")
                            {
                                OpenFile();
                            }
                             
                        }
                        break;

                    case BoEventTypes.et_KEY_DOWN:
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {

                        }
                        break;
                        //default:
                }
            }
            catch (Exception ex)
            {


            }
            finally
            {
              /*  if (oForm != null)
                    oForm.Freeze(false);*/
            }
            return BubbleEvent;
        }


        public void OpenFile()
        {
            try
            {
                System.Threading.Thread ShowFolderBrowserThread;
                ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);
                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void ShowFolderBrowser()
        {
            try
            {

                oForm = SBOMain.SBO_Application.Forms.Item(cFormID);

                NativeWindow nws = new NativeWindow();
                //System.Windows.Forms.Form nws = new Form();
                OpenFileDialog MyTest = new OpenFileDialog();
                Process[] MyProcs = null;
                string filename = null;

                MyProcs = Process.GetProcessesByName("SAP Business One");
                nws.AssignHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);

                IntPtr ptr = GetForegroundWindow();
                WindowWrapper oWindow = new WindowWrapper(ptr);

                if (MyTest.ShowDialog(oWindow) == System.Windows.Forms.DialogResult.OK)
                {
                    filename = MyTest.FileName;
                    BrowseFilePath = filename;
                    oForm.Items.Item("tattach").Specific.value = filename;
                    if (oForm.Mode == BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                    System.Windows.Forms.Application.ExitThread();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void SetCode()
        {

            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            string TableName = "EXDM";
            SBOMain.SetCode(oForm.UniqueID, TableName);

            //throw new NotImplementedException();
        }

        public void Form_Load_Components(SAPbouiCOM.Form oForm,string mode)
        {
            if (mode != "OK")
            {
                oForm.Freeze(true);
                SAPbouiCOM.OptionBtn tstatus = (SAPbouiCOM.OptionBtn)oForm.Items.Item("tstatus").Specific;
                SAPbouiCOM.OptionBtn tstatusI = (SAPbouiCOM.OptionBtn)oForm.Items.Item("tstatusI").Specific;
                tstatusI.GroupWith("tstatus");

                SetCode();
                tstatus.Selected = true;
                SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("cdoctype").Specific;
                cb.ExpandType = BoExpandType.et_DescriptionOnly;
                cb.Select("E");
                oForm.Freeze(false);
            }
        }

    }
}
