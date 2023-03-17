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
using System.Collections.Specialized;
using CoreSuteConnect.Events;
using CoreSuteConnect.Class.DEFAULTSAPFORMS;

namespace CoreSuteConnect.Class
{
    public class CommonUtility
    {
        public bool IsNullOrEmpty(SAPbouiCOM.Form oForm, string field, string msg)
        {  
            if (String.IsNullOrEmpty(oForm.Items.Item(field).Specific.Value.ToString()))
            {
                SBOMain.SBO_Application.StatusBar.SetText("Please " + msg , BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                 return false;
            }
            else
            {
                return true;
            }
        }
        public void FormLoadAndActivate(string formName, string menuId)
        {
            try
            {
                bool plFormOpen = false;
                for (int i = 1; i < SBOMain.SBO_Application.Forms.Count; i++)
                {
                    if (SBOMain.SBO_Application.Forms.Item(i).UniqueID == formName)
                    {
                        SBOMain.SBO_Application.Forms.Item(i).Select();
                        plFormOpen = true;
                    }
                }
                if (!plFormOpen)
                {
                    SBOMain.SBO_Application.Menus.Item(menuId).Activate();
                }
            }
            catch (Exception ex)
            {

            }
        }
       public double getLineTotalFromDocKey(string table, string docentry)
        {
            double lineTotal = 0;
            string strQry = "Select Sum(LineTotal) as 'lineTotal' from " + table + "  Where DocEntry = '" + docentry + "'";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(strQry);
            lineTotal = rec.Fields.Item("lineTotal").Value;
            return lineTotal;
        }

        public int getTableRecordCount(string table)
        {
            int Total = 0;
            string strQry = "Select count(*) as TotalRec from [dbo].[@" + table +"]";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(strQry);
            Total = rec.Fields.Item("TotalRec").Value;
            return Total;
        }

        public double getFCLineTotalFromDocKey(string table, string docentry)
        {
            double lineTotal = 0;
            string strQry = "Select Sum(TotalFrgn) as 'lineTotal' from " + table + "  Where DocEntry = '" + docentry + "'";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(strQry);
            lineTotal = rec.Fields.Item("lineTotal").Value;
            return lineTotal;

        }
        public double getRateFromDocKey(string table, string docentry)
        {
            double lineTotal = 0;
            string strQry = "Select Sum(TotalFrgn) as 'lineTotal' from " + table + "  Where DocEntry = '" + docentry + "'";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(strQry);
            lineTotal = rec.Fields.Item("lineTotal").Value;
            return lineTotal;

        }
        public string getDocNumFromDocKey(string table, string docentry)
        { 
            string docnum = null;
            string strQry = "Select DocNum from " + table + "  Where DocEntry = '" + docentry + "'";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(strQry);
            docnum = Convert.ToString(rec.Fields.Item("DocNum").Value);
            return docnum;
        }
        public string getCurrFromDocKey(string table, string docentry)
        {
            string docnum = null;
            string strQry = "Select DISTINCT(Currency)  as 'Currency' from " + table + "  Where DocEntry = '" + docentry + "'";
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            rec.DoQuery(strQry);
            docnum = Convert.ToString(rec.Fields.Item("Currency").Value);
            return docnum;
        }
        public void doAutoColSum(SAPbouiCOM.Matrix matrix, string ColumnName)
        {
            SAPbouiCOM.Column mCol = matrix.Columns.Item(ColumnName);
            mCol.RightJustified = true;
            mCol.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
        }

        public static void LoadFromXML(string FileName, string ModulName)
        {
            try
            {
                System.Xml.XmlDocument oXmlDoc = null;
                oXmlDoc = new System.Xml.XmlDocument();
                // load the content of the XML File
                string sPath = null;
                FileName = ModulName + @"\" + FileName + ".xml";
                sPath = System.Windows.Forms.Application.StartupPath;
                oXmlDoc.Load(sPath + @"\Forms\" + FileName);
                // load the form to the SBO application in one batch
                string tmpStr;
                tmpStr = oXmlDoc.InnerXml;
                SBOMain.SBO_Application.LoadBatchActions(ref tmpStr);
                sPath = SBOMain.SBO_Application.GetLastBatchResults();
                // oForm = SBO_Application.Forms.ActiveForm;
            }
            catch (Exception ex)
            {
                SBOMain.SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }

        }

        public DateTime Add_Year(DateTime CurrDate)
        {
            return CurrDate.AddYears(1);
        }

        public DateTime Add_Month(DateTime CurrDate)
        {
            return CurrDate.AddMonths(6);
        }

        public string BPTopOneAddress(string cardcode)
        {
            SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "SELECT TOP 1 T1.[Block], T1.[Building], T1.[Street], T1.[City], T1.[ZipCode], T2.[Name]  AS 'State', T3.[Name] as 'Country' FROM[dbo].[CRD1] AS T1 ";
            query = query + "LEFT JOIN OCST AS T2 ON T1.[State] = T2.[Code] AND T1.[Country] = T2.[Country]";
            query = query + "LEFT JOIN OCRY AS T3 ON T1.[Country] = T3.[Code]";
            query = query + "WHERE T1.[CardCode] = '" + cardcode + "' AND T1.[AdresType] = 'B'";

            string Address = null;

            rec.DoQuery(query);
             
            if (rec.RecordCount > 0)
            {
                while (!rec.EoF)
                {     
                    string block = rec.Fields.Item("Block").Value.ToString();
                    string Building = rec.Fields.Item("Building").Value.ToString();
                    string Street = rec.Fields.Item("Street").Value.ToString();
                    string City = rec.Fields.Item("City").Value.ToString();
                    string ZipCode = rec.Fields.Item("ZipCode").Value.ToString();
                    string State = rec.Fields.Item("State").Value.ToString();
                    string Country = rec.Fields.Item("Country").Value.ToString();

                    Address = block + "," + Building + ", " + Environment.NewLine;
                           Address = Address + Street + ", " + Environment.NewLine;
                           Address = Address + City + " - " + ZipCode + Environment.NewLine;
                           Address = Address + State + " - " + Country ;
                     
                    rec.MoveNext();
                }
                return Address;
            }
            else
            {
                return "";
            } 
        }

        public string BankAddress(string bankcode)
        {
            try
            { 
            
                SAPbobsCOM.Recordset rec = (SAPbobsCOM.Recordset)SBOMain.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = "SELECT T1.* FROM ODSC AS T0 LEFT JOIN DSC1 AS T1 ON T0.BankCode = T1.BankCode WHERE T1.[BankCode] = '" + bankcode+ "'";
                string Address = null;

                rec.DoQuery(query);

                if (rec.RecordCount > 0)
                {
                    while (!rec.EoF)
                    {
                        string block = rec.Fields.Item("Block").Value.ToString();
                        string Building = rec.Fields.Item("Building").Value.ToString();
                        string Street = rec.Fields.Item("Street").Value.ToString();
                        string City = rec.Fields.Item("City").Value.ToString();
                        string ZipCode = rec.Fields.Item("ZipCode").Value.ToString();
                        string State = rec.Fields.Item("State").Value.ToString();
                        string Country = rec.Fields.Item("Country").Value.ToString();

                        Address = block + "," + Building + ", " + Environment.NewLine;
                        Address = Address + Street + ", " + Environment.NewLine;
                        Address = Address + City + " - " + ZipCode + Environment.NewLine;
                        Address = Address + State + " - " + Country;

                        rec.MoveNext();
                    }
                    return Address;
                }
                else
                {
                    return "";
                }
            }catch(Exception ex)
            {
                return "";
            }
        }
         
        public int GetNextDocNum(ref SAPbouiCOM.EditText oEdit, ref string TableName)
        {
            try
            {
                oEdit.Value = string.Empty;
                if (TableName.Trim() != string.Empty)
                {
                    SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
                    oRec = SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRec.DoQuery("Select Max(\"DocEntry\") From \"" + TableName + "\"");
                    int MaxCode = Convert.ToInt32(oRec.Fields.Item(0).Value);
                    if (MaxCode == 0)
                    {
                        oEdit.Value = "1";
                    }
                    else
                    {
                        if (!oRec.EoF)
                        {
                            oEdit.Value = Convert.ToString(MaxCode + 1);
                        }
                        else
                        {
                            oEdit.Value = "1";
                        }
                    }
                    oRec = null;
                }
                else
                {
                    SBOMain.SBO_Application.StatusBar.SetText("Error on GetNextDocNum : Invalid TableName ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
                return 0;
            }
            catch (Exception ex)
            {
                SBOMain.SBO_Application.StatusBar.SetText("Error on GetNextDocNum :" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return -1;
            }
        }
    }
}
