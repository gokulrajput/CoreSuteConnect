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

namespace CoreSuteConnect.Class.Common
{
    class clsLicenceManager
    {
        #region VariableDeclaration
        private SAPbouiCOM.Form oForm;
        public string cFormID = string.Empty;

        public static string getSalesForm = string.Empty;

        CommonUtility objCU = new CommonUtility();

        #endregion VariableDeclaration

        public bool MenuEvent(ref SAPbouiCOM.MenuEvent pVal, string FormId)
        {
            bool BubbleEvent = true;
            try
            {
                cFormID = FormId;
                oForm = SBOMain.SBO_Application.Forms.Item(FormId);
                Form_Load_Components(oForm, "ADD"); 
            }
            catch (Exception ex)
            { 

            }
            finally
            {

            }

            return BubbleEvent;
        }

        public void Form_Load_Components(SAPbouiCOM.Form oForm, string mode)
        { 
            oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;
            oForm.Items.Item("tab1").Visible = true;
            oForm.Items.Item("tab1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.PaneLevel = 1;
             
            oForm.DataSources.DataTables.Add("tab");
            SAPbouiCOM.Grid objGrid = oForm.Items.Item("grdUsers").Specific;
             
            string Qry = "SELECT USER_CODE AS 'User Code', U_NAME AS 'UserName' FROM OUSR";
            oForm.DataSources.DataTables.Item("tab").ExecuteQuery(Qry);
            objGrid.DataTable = oForm.DataSources.DataTables.Item("tab");
             
            oForm.DataSources.DataTables.Add("tab5");
            SAPbouiCOM.Grid objGrid1 = oForm.Items.Item("grdModules").Specific;

            string Qry1 = "SELECT USER_CODE AS 'User Code', U_NAME AS 'UserName' FROM OUSR";
            oForm.DataSources.DataTables.Item("tab5").ExecuteQuery(Qry1);
            objGrid1.DataTable = oForm.DataSources.DataTables.Item("tab5");

           /*oForm.DataSources.DataTables.Add("MyTable");
            SAPbouiCOM.DataTable oDataTabe = oForm.DataSources.DataTables.Item("MyTable");
            oDataTabe.Clear();
            oDataTabe.Columns.Add("0", BoFieldsType.ft_AlphaNumeric);
            oDataTabe.Columns.Add("1", BoFieldsType.ft_AlphaNumeric); 
               oDataTabe.Rows.Add();
               oDataTabe.SetValue("0", 1, "EXIM");
               oDataTabe.SetValue("1", 1, "EXIM");*/
        }

        public bool ItemEvent(ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent, string FormId)
        {
            BubbleEvent = true;
            oForm = SBOMain.SBO_Application.Forms.Item(FormId);
            try
            {
                switch (pVal.EventType)
                { // KEY Down event for multiplication

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm; 
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        { 
                            
                        }
                        break;

                    case BoEventTypes.et_DOUBLE_CLICK:
                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {
                             
                        }
                        if (pVal.BeforeAction == false)
                        { 
                                
                        }
                        break;

                    case BoEventTypes.et_CLICK:

                        oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                        if (pVal.BeforeAction == true)
                        {

                        }
                        if (pVal.BeforeAction == false)
                        {

                        }

                        break;

                }
            }
            catch (Exception ex)
            {

            }

            return BubbleEvent;
        }
    }
}

