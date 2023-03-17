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
using CoreSuteConnect.Class.DEFAULTSAPFORMS;

namespace CoreSuteConnect.Class.EXIM
{
    class clsInctMaster
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;

        OpenFileDialog OpenFileDialog = new OpenFileDialog();
        string BrowseFilePath = string.Empty;

        #endregion VariableDeclaration

        public clsInctMaster(OutwardToIncoMaster inClass)
        {
            if (inClass != null)
            {
                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                oForm.Items.Item("tinctcode").Specific.value = inClass.inctno;
                oForm.Items.Item("1").Click();
            }
        }

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId, string Type = "")
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);

                if(oForm.Mode == BoFormMode.fm_ADD_MODE || Type == "ADDNEWFORM")
                {
                    Form_Load_Components(oForm, "ADD");
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
                            oForm = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        if (pVal.BeforeAction == false)
                        {

                        } 
                        break;
                    case BoEventTypes.et_CLICK:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm; 
                        if(pVal.BeforeAction == true)
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
                                string FileName = Path.GetFileName(oForm.Items.Item("tattach").Specific.value);
                                string ReplaceFilePath = SBOMain.Get_Attach_Folder_Path() + FileName;
                                File.Move(BrowseFilePath, ReplaceFilePath);
 
                            }
                        } 

                        break;
                    case BoEventTypes.et_ITEM_PRESSED:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;

                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "1" && (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE))
                            {
                                if (string.IsNullOrEmpty(oForm.Items.Item("tinctcode").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Incoterm Code", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tinctcode").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("tinctname").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Incoterm Name", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("tinctname").Click();
                                } 
                            }
                        }
                        if(pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1" && oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                Form_Load_Components(oForm,"ADD");
                            }
                            else if (pVal.ItemUID == "btnATC")
                            {
                                OpenFile();
                            }
                             
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
                /*if (oForm != null)
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
            string TableName = "EXIM";
            SBOMain.SetCode(oForm.UniqueID, TableName);
            
        }
        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        {
            if (mode != "OK")
            {
                oForm.Freeze(true);
                SAPbouiCOM.OptionBtn tstatus = (SAPbouiCOM.OptionBtn)oForm.Items.Item("tstatus").Specific;
                SAPbouiCOM.OptionBtn tstatusI = (SAPbouiCOM.OptionBtn)oForm.Items.Item("tstatusI").Specific;
                tstatusI.GroupWith("tstatus");

                SetCode();
                tstatus.Selected = true;

                oForm.Freeze(false);
            }
        }
    }
}
