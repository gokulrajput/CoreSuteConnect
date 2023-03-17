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
    class clsExpMaster
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;
        private SAPbouiCOM.ComboBox oComboBox;

        private SAPbouiCOM.ChooseFromList oCFL = null;
        private SAPbouiCOM.Conditions oConds = null;
        private SAPbouiCOM.Condition oCond = null;

        OpenFileDialog OpenFileDialog = new OpenFileDialog();
        string BrowseFilePath = string.Empty;

        #endregion VariableDeclaration

        public clsExpMaster(OutwardToEXPMaster inClass)
        {
            if (inClass != null)
            {
                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                oForm.Items.Item("ItemCode").Specific.value = inClass.itemcode;
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
                if (oForm.Mode == BoFormMode.fm_ADD_MODE || Type == "ADDNEWFORM")
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
                            if (pVal.ItemUID == "ItemCode")
                            {
                                CFLCondition("cfl1");
                            }
                        }
                        if (pVal.BeforeAction == false)
                        { 
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                            string sCFL_ID = oCFLEvento.ChooseFromListUID;
                            SAPbouiCOM.ChooseFromList oCFL = null;
                            oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            if (oDataTable != null)
                            {
                                if (pVal.ItemUID == "ItemCode")
                                { 
                                    try
                                    {
                                        oForm.Items.Item("ItemCode").Specific.value = oDataTable.GetValue("ItemCode", 0).ToString();
                                        oForm.Items.Item("texpname").Specific.value = oDataTable.GetValue("ItemName", 0).ToString();
                                        oForm.Items.Item("texpadnm").Specific.value = oDataTable.GetValue("FrgnName", 0).ToString();
                                    }
                                    catch (Exception ex)
                                    { 
                                    }

                                }
                            }
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
                                if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Expense Code", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("ItemCode").Click();
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("texpname").Specific.value))
                                {
                                    BubbleEvent = false;
                                    SBOMain.SBO_Application.StatusBar.SetText("Please Add Expense Name", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    oForm.Items.Item("texpname").Click();
                                }
                                else
                                {
                                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cexptype").Specific;
                                    if (oComboBox.Selected == null)
                                    {
                                        BubbleEvent = false;
                                        SBOMain.SBO_Application.StatusBar.SetText("Please Select Expense Type", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        oForm.Items.Item("cexptype").Click();
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
            string TableName = "EXEM";
            SBOMain.SetCode(oForm.UniqueID, TableName);
            
            //throw new NotImplementedException();
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
                SAPbouiCOM.ComboBox cb = (SAPbouiCOM.ComboBox)oForm.Items.Item("cexptype").Specific;
                cb.ExpandType = BoExpandType.et_DescriptionOnly;
                cb.Select("E");
                oForm.Freeze(false);
            }
        }
        private void CFLCondition(string CFLID)
        {
            oCFL = oForm.ChooseFromLists.Item(CFLID);
            oConds = SBOMain.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
            if (CFLID == "cfl1")
            {
                oCond = oConds.Add();
                oCond.Alias = "PrchseItem";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";
                   
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "SellItem";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "InvntItem";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";

                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "ItemClass";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "1";
   
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND; 
                oCond = oConds.Add();
                oCond.Alias = "validFor";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y"; 

                oCFL.SetConditions(oConds);

            } 
            oCFL = null;
            oCond = null;
            oConds = null;

        }
    }
}
