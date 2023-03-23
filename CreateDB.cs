using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using CoreSuteConnect.Class.EXIM;
using CoreSuteConnect.Class.PRICELIST;
using CoreSuteConnect.Class.DEFAULTSAPFORMS;
using CoreSuteConnect.Class.QC;
//using CoreSuteConnect.Class.QC.QCDB;

namespace CoreSuteConnect
{    
    internal class CreateDB
    {
        #region variableDeclaration
        private SAPbobsCOM.UserTablesMD oUserTablesMD = null;
        private SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
        private SAPbobsCOM.Recordset oRecordset;
        private SAPbobsCOM.UserObjectsMD oUserObjectMD = null;

        //QCDB QCDB = new QCDB();

        #endregion
        public void createDB()
        {
            if (CreateDataBase() == true)
            {
                //System.Windows.Forms.MessageBox.Show("Add-On database tables created successfully!");
            }
        }
        private bool CreateDataBase()
        {
            try
            {
               // QCDB.createDB();

                CreateTable("QCSR", "QC Sample Request", SAPbobsCOM.BoUTBTableType.bott_Document);
                FieldDetails("@QCSR", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                FieldDetails("@QCSR", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@QCSR", "Dscription", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@QCSR", "CardCode", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@QCSR", "CardName", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@QCSR", "DocDate", "Doc Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@QCSR", "BatchNo", "Batch No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@QCSR", "BaseDocNo", "Base DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@QCSR", "BaseDocEnt", "Base DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@QCSR", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@QCSR", "RefNo", "Reference No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@QCSR", "CTime", "Created Time", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, 100, "", false, "");
                FieldDetails("@QCSR", "BrcNm", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@QCSR", "Whs", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@QCSR", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 100, "", false, "");
                FieldDetails("@QCSR", "ItemGrp", "Item Group", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@QCSR", "InOutNo", "Inward/Outward No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");

                 
                CreateTable("CSR1", "QC Sample Data", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@CSR1", "ParamCode", "Parameter Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@CSR1", "ParamName", "Parameter Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@CSR1", "Method", "Method", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@CSR1", "StdReslt", "Standard Result", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@CSR1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@CSR1", "Strength", "Strength", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@CSR1", "Attach", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 0, "", false, "");
                FieldDetails("@CSR1", "DL", "DL", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@CSR1", "DA", "DA", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@CSR1", "DE", "DE", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@CSR1", "DB", "DB", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@CSR1", "DC", "DC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@CSR1", "DH", "DH", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@CSR1", "DC", "DC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                //FieldDetails("@CSR1", "InOutNo", "Inward/Outward No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 50, "", false, "");
               // FieldDetails("@CSR1", "PRNo", "Party Ref no", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 50, "", false, "");


                List<string> FindColumnQCSR = new List<string>();
                List<string> ChildTableQCSR = new List<string>();
                string[] DatalistChildTableQCSR = { "CSR1"};
                ChildTableQCSR.AddRange(DatalistChildTableQCSR);
                string[] DatalistfcolQCSR = { "U_DocType", "U_ItemCode", "U_Dscription", "U_CardName", 
                    "U_CardCode", "U_DocDate", "U_BatchNo", "U_BaseDocNo", "U_BaseDocEnt",  "U_RefNo", "U_InOutNo"};

                FindColumnQCSR.AddRange(DatalistfcolQCSR); 
                CreateUserObject("QCSR", "QC Sample Request", "QCSR", ChildTableQCSR, FindColumnQCSR, BoYesNoEnum.tYES, BoUDOObjType.boud_Document);
                  
                ////////////////////////////
                // **** POST MASTER TAMBLE *****88
                CreateTable("EXPM", "PortMaster", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                FieldDetails("@EXPM", "portcode", "Port Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXPM", "portname", "Port Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXPM", "sspno", "Self Sealing Permission No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXPM", "attachment", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 250, "", false, "");
                FieldDetails("@EXPM", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXPM", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                
                List<string> FindColumnPM = new List<string>();
                List<string> ChildTablePM = new List<string>();
                string[] DatalistPM = { "U_portcode", "U_portname","U_remarks", "U_status", "U_sspno" };
                FindColumnPM.AddRange(DatalistPM); 
                CreateUserObject("EXPM", "Port Master", "EXPM", ChildTablePM, FindColumnPM, BoYesNoEnum.tNO, BoUDOObjType.boud_MasterData);
                 
                // **** INCOTERM MASTER TAMBLE *****
                CreateTable("EXIM", "IncotermMaster", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                FieldDetails("@EXIM", "inctcode", "Incoterm Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXIM", "inctname", "Incoterm Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXIM", "attachment", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 0, "", false, "");
                FieldDetails("@EXIM", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXIM", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                List<string> FindColumnIM = new List<string>();
                List<string> ChildTableIM = new List<string>();
                string[] DatalistIM = { "U_inctcode", "U_inctname", "U_remarks", "U_status"};
                FindColumnIM.AddRange(DatalistIM);
                CreateUserObject("EXIM", "Incoterm Master", "EXIM", ChildTableIM, FindColumnIM, BoYesNoEnum.tNO, BoUDOObjType.boud_MasterData);

                // **** DOCUMENT MASTER TAMBLE *****
                CreateTable("EXDM", "DocumentMaster", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                FieldDetails("@EXDM", "doctype", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXDM", "doccode", "Document Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXDM", "docname", "Document Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXDM", "attachment", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 0, "", false, "");
                FieldDetails("@EXDM", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXDM", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                List<string> FindColumnDOC = new List<string>();
                List<string> ChildTableDOC = new List<string>();
                string[] DatalistDOC = { "U_doctype", "U_doccode", "U_docname", "U_remarks" , "U_status" };
                FindColumnDOC.AddRange(DatalistDOC);                 
                CreateUserObject("EXDM", "Document Master", "EXDM", ChildTableDOC, FindColumnDOC, BoYesNoEnum.tNO, BoUDOObjType.boud_MasterData);
                 
                CreateTable("EXEM", "ExpenseMaster", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                FieldDetails("@EXEM", "exptype", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXEM", "expcode", "Expense Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXEM", "expname", "Expense Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@EXEM", "expadname", "Expence Additional Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@EXEM", "attachment", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 0, "", false, "");
                FieldDetails("@EXEM", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXEM", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                List<string> FindColumnEM = new List<string>();
                List<string> ChildTableEM = new List<string>();
                string[] DatalistEM = { "U_exptype", "U_expcode", "U_expname", "U_expadname", "U_remarks", "U_status" };
                FindColumnEM.AddRange(DatalistEM); 
                CreateUserObject("EXEM", "Expense Master", "EXEM", ChildTableEM, FindColumnEM, BoYesNoEnum.tNO, BoUDOObjType.boud_MasterData);
                  
                
                CreateTable("EXSM", "Scheme Master", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                FieldDetails("@EXSM", "schtype", "Scheme Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@EXSM", "schno", "Scheme No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "schname", "Scheme Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXSM", "schrate", "Scheme Rate (Percentage)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Rate, 20, "", false, "");
                FieldDetails("@EXSM", "schratepk", "Scheme Rate (Per KG)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@EXSM", "schsd", "Scheme Start Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0 ,"", false, "");
                FieldDetails("@EXSM", "sched", "Scheme End Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXSM", "schLEDI", "License Expiry Date for Import", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXSM", "schLEDE", "License Expiry Date for Export", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXSM", "schTPLP", "Third Party Lincese Purchase", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXSM", "schvc", "Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "schvn", "Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXSM", "schapin", "A/P Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "schapinde", "A/P Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "schIA", "Issuing Authority", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "schSOF", "Submission of File", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None,0, "", false, "");
                FieldDetails("@EXSM", "DGFTRD", "DGFT Rec. Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXSM", "DGFTFN", "DGFT File No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "PRN", "Port Reg. No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "PRD", "Port Reg. Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXSM", "custVer", "Custom Verification", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXSM", "ext", "Extension", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@EXSM", "TPLS", "Third Party Lincese Sales", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXSM", "custcd", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "custnm", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXSM", "arinv", "A/R Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "arinvde", "A/R Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXSM", "attachment", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 0, "", false, "");
                FieldDetails("@EXSM", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXSM", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                 
                CreateTable("XSM1", "Export Obligation", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@XSM1", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM1", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XSM1", "qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@XSM1", "uom", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@XSM1", "amtLC", "Amount (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM1", "amtFC", "Amount (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM1", "flqty", "Fulfilment Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@XSM1", "fllC", "Fulfilment (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM1", "flfC", "Fulfilment (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM1", "rmqty", "Remaining Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity,20, "", false, "");
                FieldDetails("@XSM1", "rmLC", "Remaining (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM1", "rmFC", "Remaining (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");

                CreateTable("XSM2", "Export Fulfilment", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@XSM2", "expexmno", "Exim Transaction No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM2", "expinvno", "Export Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None,50, "", false, "");
                FieldDetails("@XSM2", "expinvde", "Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM2", "expinvdt", "Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XSM2", "expbn", "Shipping Bill No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM2", "expbd", "Shipping Bill Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XSM2", "expfq", "Fulfilment Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 50, "", false, "");
                FieldDetails("@XSM2", "expuom", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@XSM2", "expfvFC", "Fulfilement Value (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM2", "expfvLC", "Fulfilement Value (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");

                CreateTable("XSM3", "Import Rights", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@XSM3", "iritemcd", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM3", "iritemnm", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XSM3", "irqty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");                
                FieldDetails("@XSM3", "iruom", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@XSM3", "irQtyPer", "Quantity %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@XSM3", "irExQtP", "Exempt Qty %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@XSM3", "iramtLC", "Amount (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM3", "iramtFC", "Amount (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM3", "irutqty", "Utilized Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");              
                FieldDetails("@XSM3", "irutlC", "Utilized (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM3", "irutfC", "Utilized (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM3", "irrmqty", "Remaining Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@XSM3", "irrmLC", "Remaining (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XSM3", "irrmFC", "Remaining (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");

                CreateTable("XSM4", "Import Utilization", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@XSM4", "impinvno", "Import Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM4", "impinvdt", "Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XSM4", "impbedn", "Bill of Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM4", "impbedt", "Bill of Entry Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XSM4", "impq", "Utilized Quantity", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM4", "impuom", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@XSM4", "impvFC", "Utilized Value (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 50, "", false, "");
                FieldDetails("@XSM4", "impvLC", "Utilized Value (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");

                CreateTable("XSM5", "HSN Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@XSM5", "hsncode", "HSN Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XSM5", "hsnname", "HSN Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");
              
                CreateTable("XSM6", "Attachments", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@XSM6", "trgtpath", "Target Path", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");
                FieldDetails("@XSM6", "filename", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");
                FieldDetails("@XSM6", "atchdate", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XSM6", "fretext", "Free Text", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XSM6", "cpytotd", "Copy To Target Document", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");
                  
                List<string> FindColumn = new List<string>();
                List<string> ChildTable = new List<string>();
                string[] DatalistChildTable = { "XSM1", "XSM2", "XSM3", "XSM4", "XSM5", "XSM6" };
                ChildTable.AddRange(DatalistChildTable); 
                string[] Datalistfcol = { "U_schtype", "U_schno", "U_schname", "U_schrate", "U_schratepk", "U_schsd", "U_sched", "U_schLEDI" , "U_schLEDE",
                                          "U_schTPLP","U_schvc","U_schvn","U_schapin","U_schapinde", "U_schIA", "U_schSOF","U_DGFTRD","U_DGFTFN","U_PRN",
                                          "U_PRD","U_custVer","U_ext","U_TPLS", "U_custcd","U_custnm","U_arinv","U_arinvde", "U_remarks", "U_status" };
                FindColumn.AddRange(Datalistfcol);
                CreateUserObject("EXSM", "Scheme Master", "EXSM", ChildTable, FindColumn, BoYesNoEnum.tYES, BoUDOObjType.boud_MasterData);
                 

                CreateTable("EXRU", "Scheme Transection", SAPbobsCOM.BoUTBTableType.bott_Document);
                FieldDetails("@EXRU", "schsn", "Script No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXRU", "schaa", "Scrip Applied Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
                FieldDetails("@EXRU", "schra", "Scrip Received Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 100, "", false, "");
                FieldDetails("@EXRU", "schup", "Utilized %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Rate, 10, "", false, "");
                FieldDetails("@EXRU", "schua", "Scrip Utilized Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 50, "", false, "");
                FieldDetails("@EXRU", "schrma", "Scrip Remaining Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 50, "", false, "");
                FieldDetails("@EXRU", "schsd", "Scheme Start Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXRU", "sched", "Expity Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXRU", "schari", "A/R Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, ""); 
                FieldDetails("@EXRU", "schrm", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXRU", "schst", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXRU", "schtype", "Scheme Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");

                CreateTable("XRU1", "Details of Export Transection", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XRU1", "schexno", "Exim Tracking No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XRU1", "schsrno", "SR No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@XRU1", "schpcpn", "Port Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None,50, "", false, "");
                FieldDetails("@XRU1", "schinvno", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XRU1", "schinvde", "Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XRU1", "schinvdt", "Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XRU1", "schblno", "Shipping Bill No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@XRU1", "schbldt", "Shipping Bill Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XRU1", "schfob", "FOB Value", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XRU1", "schapamt", "Applied Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XRU1", "schrcamt", "Received Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                 
                CreateTable("XRU2", "Details of Import Transection", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XRU2", "impexno", "Exim Tracking No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XRU2", "impsrno", "Sr No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@XRU2", "impinvno", "Import Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XRU2", "impinvde", "Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XRU2", "impinvdt", "Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XRU2", "impbeno", "Bill of Entry No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XRU2", "impbedt", "Bill of Entry Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XRU2", "impAV", "Assessable Value", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XRU2", "impUA", "Utilized Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");

                List<string> FindColumnST = new List<string>();
                List<string> ChildTableST = new List<string>(); 
                string[] DatalistChildTable1 = { "XRU1", "XRU2" };
                ChildTableST.AddRange(DatalistChildTable1); 
                string[] Datalistfcol1 = { "U_schsn", "U_schaa", "U_schra", "U_schup", "U_schua", "U_schrma", "U_schsd", "U_sched", "U_schari",  "U_schrm", "U_schst", "U_schtype" };
                FindColumnST.AddRange(Datalistfcol1); 
                CreateUserObject("EXRU", "Scheme Transection", "EXRU", ChildTableST, FindColumnST, BoYesNoEnum.tYES, BoUDOObjType.boud_Document);
                  
                CreateTable("EXET", "EXIM Tracking", SAPbobsCOM.BoUTBTableType.bott_Document);
                FieldDetails("@EXET", "exdt", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXET", "exsn", "Serial No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@EXET", "exbc", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exbn", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXET", "exson", "Sales Order No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exsonde", "Sales Order DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exdn", "Delivery No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exdnde", "Delivery DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exinvno", "A/R Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exinvnode", "A/R Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exinvdt", "A/R Invoice Ex. Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Rate, 0, "", false, "");
                FieldDetails("@EXET", "exlcno", "LC Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exlcdt", "LC Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");                
                FieldDetails("@EXET", "exdd", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXET", "exinco", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXET", "expcgb", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXET", "expcrb", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXET", "exsbn", "Shipping Bill Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "exsbd", "Shipping Bill Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXET", "exls", "Licensed / Scheme", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@EXET", "exprm", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXET", "exst", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXET", "twbwc", "Weighing bridge Vendor Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "twbwn", "Weighing bridge Vendor Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXET", "twbwa", "Weighing bridge Vendor Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("@EXET", "twsn", "Weighing Slip No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXET", "twbn", "Booking No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXET", "twbd", "Weighing bridge Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXET", "twbt", "Weighing bridge Time", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", false, "");
                FieldDetails("@EXET", "twpl", "Pallets", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXET", "twpn", "No of Pallets", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                

                CreateTable("XET1", "Shipping Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET1", "ex1pol", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET1", "ex1pod", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET1", "ex1por", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET1", "ex1fd", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET1", "ex1coo", "Country of Origin", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET1", "ex1dc", "Destination Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET1", "ex1eld", "ETD Loading Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET1", "ex1edd", "ETA Discharge Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET1", "ex1efdd", "ETA Final Destin. Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET1", "ex1atd", "Approx. Transit Days", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                FieldDetails("@XET1", "ex1icn", "Insurance Cover No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET1", "ex1fcn", "Fumigation Certificate No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET1", "ex1cdn", "Courier Docket No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET1", "ex1bln1", "BL No - 1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET1", "ex1bld1", "BL Date - 1", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET1", "ex1bln2", "BL No - 2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET1", "ex1bld2", "BL Date - 2", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET1", "ex1vnm", "Vessal Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET1", "ex1vno", "Vessal Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET1", "ex1ald", "ATD Loading Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET1", "ex1add", "ATA Discharge Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET1", "ex1afdd", "ATA Final Destin. Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET1", "ex1actd", "Actual Transit Days", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                FieldDetails("@XET1", "ex1slc", "Shipping Liner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET1", "ex1sln", "Shipping Liner Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET1", "exrrno", "RR Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET1", "exrrdt", "RR Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");


                CreateTable("XET2", "FOB calculation", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET2", "cifAFC", "CIF Actual FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "cifALC", "CIF Actual LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "cifDFC", "CIF Display FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "cifDLC", "CIF Display LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");                 
                FieldDetails("@XET2", "frtAFC", "Fright Actual FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "frtALC", "Fright Actual LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "frtDFC", "Fright Display FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "frtDLC", "Fright Display LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");                 
                FieldDetails("@XET2", "insAFC", "Insurance Actual FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "insALC", "Insurance Actual LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "insDFC", "Insurance Display FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "insDLC", "Insurance Display LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");                 
                FieldDetails("@XET2", "ocAFC", "Other Charges Actual FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "ocALC", "Other Charges Actual LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "ocDFC", "Other Charges Display FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "ocDLC", "Other Charges Display LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "fobAFC", "FOB Value Actual FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "fobALC", "FOB Value Actual LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "fobDFC", "FOB Value Display FC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET2", "fobDLC", "FOB Value Display LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, ""); 
                FieldDetails("@XET2", "fcCHAcd", "CHA  Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET2", "fcCHAnm", "CHA Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@XET2", "fcCHApn", "CHA Phone", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Phone, 100, "", false, "");
                FieldDetails("@XET2", "fcCHAad", "CHA Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 100, "", false, "");
                FieldDetails("@XET2", "fcebrcno", "E - BRC No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET2", "fcebrcdt", "E - BRC Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET2", "fcebrcamt", "E - BRC Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                  
                CreateTable("XET3", "Expense Planned / Actual", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET3", "et3expt", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@XET3", "et3expnm", "Expense Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XET3", "et3yc", "You Can", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET3", "et3ede", "Docentry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET3", "et3edn", "Docnum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET3", "et3pf", "Planned (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
                FieldDetails("@XET3", "et3pl", "Planned (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
                FieldDetails("@XET3", "et3af", "Actual (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
                FieldDetails("@XET3", "et3al", "Actual (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
                FieldDetails("@XET3", "et3df", "Diffrence (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
                FieldDetails("@XET3", "et3dl", "Diffrence (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
                FieldDetails("@XET3", "et3de", "Detailed Expense", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@XET3", "et3cur", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                FieldDetails("@XET3", "et3rt", "Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Rate, 20, "", false, "");

                CreateTable("XET4", "Remaining License Fulfilment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET4", "ex4ic", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET4", "ex4in", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XET4", "ex4ln", "License No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET4", "ex4lv", "License Validity", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET4", "ex4lrqty", "Lic. Remain. (Qty)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET4", "ex4lrafc", "Lic. Remain. Amt (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET4", "ex4lraklc", "Lic. Remain. Amt (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET4", "ex4lfqty", "Lic. Fulfilment (Qty)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET4", "ex4lffc", "Lic. Fulfilment (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET4", "ex4lflc", "Lic. Fulfilment (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");

                CreateTable("XET5", "Raw Material Consumption", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET5", "ex5ic", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET5", "ex5in", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XET5", "ex5np", "Net %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage,0, "", false, "");
                FieldDetails("@XET5", "ex5nw", "Net Weight", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET5", "ex5exp", "Exempted %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 0, "", false, "");
                FieldDetails("@XET5", "ex5exw", "Exempted Weight", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET5", "ex5ln", "Licence No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");

                CreateTable("XET6", "DBK Scheme", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET6", "dbkTJE", "DBK Transection Journal Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET6", "dbkAJE", "DBK Actual Journal Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET6", "dbkno", "DBK No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET6", "received", "Received", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 50, "", false, "");
                 
                CreateTable("XET7", "VGM Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET7", "ex7cs", "Container Size", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET7", "ex7cn", "Container No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET7", "ex7slsn", "Shipping Line Seal No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET7", "ex7rsn", "RFID Seal No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET7", "ex7woc", "Weight of Container", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET7", "ex7cwnw", "Cargo Weight / Net Weight", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET7", "ex7pw", "Packing Weight", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET7", "ex7tw", "Tare Weight", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET7", "ex7anop", "Annuexure No of Packages", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                FieldDetails("@XET7", "ex7agw", "Annexure Gross Weight", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
                 
                CreateTable("XET8", "Container Allotment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET8", "ex8cn", "Container No",  SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET8", "ex8ic", "Item Code",     SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET8", "ex8inm", "Item Name",    SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XET8", "ex8pt", "Package Type",  SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET8", "ex8nop", "No of Packages", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None,10, "", false, "");
                FieldDetails("@XET8", "ex8qty", "Quantity",        SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@XET8", "ex8aq", "Alternate Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@XET8", "ex8tw", "Tare Weight", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@XET8", "ex8gw",  "Gross Weight", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@XET8", "ex8rm", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");

                CreateTable("XET9", "Payment Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET9", "pdipn", "Incoming Payment No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET9", "pddt", "Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@XET9", "pdamtfc", "Amount (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET9", "pdamtexr", "Amount (Ex. Rate)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Rate, 0, "", false, "");
                FieldDetails("@XET9", "pdamtlc", "Amount (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                 
                CreateTable("XET10", "DBK Scheme Calculation", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET10", "ex10ic", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET10", "ex10in", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XET10", "ex10sc", "Scheme Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET10", "ex10sn", "Scheme Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XET10", "ex10qt", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 50, "", false, "");
                FieldDetails("@XET10", "ex10fofc", "FOB (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10folc", "FOB (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10frfc", "Freight (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10frlc", "Freight (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10infc", "Insurance (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10inlc", "Insurance (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10fbv", "FOB Value", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10scp", "Scheme %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 0, "", false, "");
                FieldDetails("@XET10", "ex10sv", "Scheme Value By %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10srpk", "Scheme Rate / KG", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 0, "", false, "");
                FieldDetails("@XET10", "ex10vrpk", "Scheme Value By Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10fv", "Final Value", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET10", "ex10li", "Invoice LI", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 5, "", false, "");
                FieldDetails("@XET10", "ex10vo", "Invoice VO", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 5, "", false, "");

                CreateTable("XET11", "RoDTEP Scheme Calculation", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET11", "ex11ic", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET11", "ex11in", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XET11", "ex11sc", "Scheme Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET11", "ex11sn", "Scheme Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, ""); 
                FieldDetails("@XET11", "ex11qt", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 50, "", false, "");
                FieldDetails("@XET11", "ex11fofc", "FOB (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11folc", "FOB (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11frfc", "Freight (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11frlc", "Freight (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11infc", "Insurance (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11inlc", "Insurance (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11fbv", "FOB Value", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11scp", "Scheme %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 20, "", false, "");
                FieldDetails("@XET11", "ex11sv", "Scheme Value By %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11srpk", "Scheme Rate / KG", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@XET11", "ex11vrpk", "Scheme Value By Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11fv", "Final Value", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET11", "ex11li", "Invoice LI", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 5, "", false, "");
                FieldDetails("@XET11", "ex11vo", "Invoice VO", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 5, "", false, "");


                CreateTable("XET12", "RoDTEP Scheme", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET12", "rodTJE", "RoDTEP Transection Journal Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET12", "rodAJE", "RoDTEP Actual Journal Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET12", "rodno", "RoDTEP no", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET12", "received", "Received", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                 
                CreateTable("XET13", "RoDTEP Script", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@XET13", "ex13st", "Script Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@XET13", "ex13sn", "Script No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XET13", "ex13up", "Utilized %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 50, "", false, "");
                FieldDetails("@XET13", "ex13bv", "Basic Value", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET13", "ex13ra", "Remaining Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
                FieldDetails("@XET13", "ex13sua", "Script Utilized Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");


                List<string> FindColumnET = new List<string>();
                List<string> ChildTableET = new List<string>();
                string[] DatalistChildTable2 = { "XET1", "XET2" , "XET3", "XET4" , "XET5", "XET6" , "XET7", "XET8", "XET9", "XET10", "XET11", "XET12" , "XET13" };
                ChildTableET.AddRange(DatalistChildTable2);
                string[] Datalistfcol2 = { "U_exdt", "U_exsn", "U_exbc" , "U_exbn", "U_exson", "U_exsonde", "U_exdn", "U_exdnde", "U_exinvno", "U_exinvnode",
                                           "U_exinvdt", "U_exlcno", "U_exlcdt", "U_exdd", "U_exinco", "U_expcgb" , "U_expcrb", "U_exsbn", 
                                            "U_exsbd", "U_exls", "U_exprm", "U_exst"};
                FindColumnET.AddRange(Datalistfcol2); 
                
                CreateUserObject("EXET", "EXIM Transection", "EXET", ChildTableET, FindColumnET, BoYesNoEnum.tYES, BoUDOObjType.boud_Document);
                 
                CreateTable("EXLR", "Letter Of Credit", SAPbobsCOM.BoUTBTableType.bott_Document);
                FieldDetails("@EXLR", "lctype", "LC Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXLR", "lcln", "LC No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXLR", "lcod", "Opening Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXLR", "lcad", "Arrival Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXLR", "lced", "Expiry Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXLR", "lcnd", "Negotiation Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXLR", "lcpoe", "Place of Expiry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXLR", "lcpon", "Place of Negotiation", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@EXLR", "lcsta", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EXLR", "lcsd", "Shipment Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXLR", "lcpedt", "Presention Expiry Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EXLR", "lcnod", "No. of Days", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                FieldDetails("@EXLR", "lcfc", "From Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXLR", "lcfv", "From Vendor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EXLR", "lcrm", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                  
               CreateTable("XLR1", "LC Terms", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
               FieldDetails("@XLR1", "lc1c", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
               FieldDetails("@XLR1", "lc1amt", "LC Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR1", "lc1exrt", "LC Ex. Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR1", "lc1aulc", "Amount Utilized (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR1", "lc1aufc", "Amount Utilized (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR1", "lc1arlc", "Amount Remaining (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR1", "lc1arfc", "Amount Remaining (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR1", "lc1tlqp", "Tolerance Qty %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 20, "", false, "");
               FieldDetails("@XLR1", "lc1tlvp", "Tolerance Value %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 20, "", false, "");
               FieldDetails("@XLR1", "lc1Abc", "Bank Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1Abn", "Bank Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1Aban", "Bank Account No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1Abb", "Bank Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1Aswc", "SWIFT Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1Aba", "Bank Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 100, "", false, "");
               FieldDetails("@XLR1", "lc1ibc", "Bank Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1ibn", "Bank Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1iban", "Bank Account No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1ibb", "Bank Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1iswc", "SWIFT Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1iba", "Bank Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bbc", "Bank Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bbn", "Bank Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bban", "Bank Account No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bbb", "Bank Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bswc", "SWIFT Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bba", "Bank Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bt", "BG Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
               FieldDetails("@XLR1", "lc1bc", "Bank Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bn", "Bank Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1ban", "Bank Account No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1bamt", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
               FieldDetails("@XLR1", "lc1bext", "Ex. Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 10, "", false, "");
               FieldDetails("@XLR1", "lc1bbrn", "BG Bank Ref. No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR1", "lc1dt", "Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
               FieldDetails("@XLR1", "lc1dd", "Due Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
               FieldDetails("@XLR1", "lc1ext", "Extended", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
               
               CreateTable("XLR2", "Linked Documents", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
               FieldDetails("@XLR2", "lc2dsn", "SO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XLR2", "lc2dde", "Doc Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@XLR2", "lc2dt", "Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
               FieldDetails("@XLR2", "lc2bc", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
               FieldDetails("@XLR2", "lc2bn", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR2", "lc2brn", "BP Ref. No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR2", "lc2cr", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
               FieldDetails("@XLR2", "lc2dot", "Doc Total", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum,  0, "", false, "");
               FieldDetails("@XLR2", "lc2am", "Applied Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 0, "", false, "");
               FieldDetails("@XLR2", "lc2dq", "Doc Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 0, "", false, "");
               FieldDetails("@XLR2", "lc2aq", "Applied Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity,0, "", false, "");

               CreateTable("XLR3", "Shipment Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
               FieldDetails("@XLR3", "lc3pol", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR3", "lc3pod", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR3", "lc3fd", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR3", "lc3oc", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR3", "lc3fdc", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR3", "lc3slc", "Shipping Liner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR3", "lc3sln", "Shipping Liner Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR3", "lc3all", "Allowed", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
               FieldDetails("@XLR3", "lc3nall", "Not Allowed", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
               FieldDetails("@XLR3", "lc3pslc", "Shipping Liner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
               FieldDetails("@XLR3", "lc3psln", "Shipping Liner Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");

               CreateTable("XLR4", "Documents", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
               FieldDetails("@XLR4", "lc4doc", "Document", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
               FieldDetails("@XLR4", "lc4docnm", "Document Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
               FieldDetails("@XLR4", "lc4org", "Original", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
               FieldDetails("@XLR4", "lc4cop", "Copies", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
               FieldDetails("@XLR4", "lc4tnoc", "Total No of Copies", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
               FieldDetails("@XLR4", "lc4atc", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 0, "", false, "");

               CreateTable("XLR5", "Expenses", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
               FieldDetails("@XLR5", "lc5expt", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@XLR5", "lc5expnm", "Expense name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@XLR5", "lc5yc", "You Can", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");                
               FieldDetails("@XLR5", "lc5bdn", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
               FieldDetails("@XLR5", "lc5bden", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
               FieldDetails("@XLR5", "lc5pf", "Planned (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR5", "lc5pl", "Planned (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR5", "lc5af", "Actual (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR5", "lc5al", "Actual (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR5", "lc5df", "Diffrence (FC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR5", "lc5dl", "Diffrence (LC)", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
               FieldDetails("@XLR5", "lc5de", "Detailed Expense", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
               FieldDetails("@XLR5", "lc5cur", "Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
               FieldDetails("@XLR5", "lc5rt", "Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Rate, 0, "", false, "");
                
                CreateTable("XLR6", "Amendment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
               FieldDetails("@XLR6", "lc6amdn", "Amendment No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
               FieldDetails("@XLR6", "lc6amdf", "Amendment Field", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR6", "lc6ov", "Old Value", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR6", "lc6nv", "New Value", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
               FieldDetails("@XLR6", "lc6atc", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 0, "", false, "");
               FieldDetails("@XLR6", "lc6rm1", "Remarks 1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");
               FieldDetails("@XLR6", "lc6rm2", "Remarks 2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");

               CreateTable("XLR7", "Attachment", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
               FieldDetails("@XLR7", "lc7tp", "Target Path", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");
               FieldDetails("@XLR7", "lc7fn", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");
               FieldDetails("@XLR7", "lc7ad", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
               FieldDetails("@XLR7", "lc7ft", "Free Text", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
               FieldDetails("@XLR7", "lc7cttd", "Copy to Target Document", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", false, "");

               List<string> FindColumnLC = new List<string>();
               List<string> ChildTableLC = new List<string>();
               string[] DatalistChildTable3 = { "XLR1", "XLR2", "XLR3", "XLR4", "XLR5", "XLR6", "XLR7"};
               ChildTableLC.AddRange(DatalistChildTable3);
               string[] Datalistfcol3 = { "U_lctype", "U_lcln", "U_lcod", "U_lcad", "U_lced", "U_lcnd", "U_lcpoe", "U_lcpon", "U_lcsta", "U_lcsd", "U_lcpedt", "U_lcnod", "U_lcfc", "U_lcfv", "U_lcrm" };
               FindColumnLC.AddRange(Datalistfcol3);
               CreateUserObject("EXLR", "Letter Of Credit", "EXLR", ChildTableLC, FindColumnLC, BoYesNoEnum.tYES, BoUDOObjType.boud_Document); 

                CreateTable("XDOC", "Expenses doc", SAPbobsCOM.BoUTBTableType.bott_Document);
                List<string> FindColumnDM = new List<string>();
                List<string> ChildTableDM = new List<string>();

                CreateTable("DOC1", "Detailed Expense", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@DOC1", "exinvsn", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@DOC1", "exinvdt", "Invoice Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@DOC1", "exbc", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@DOC1", "exbn", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@DOC1", "exvrn", "Vendor Ref. No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@DOC1", "excurr", "Curr.", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                FieldDetails("@DOC1", "exrate", "Ex. Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Rate, 10, "", false, "");
                FieldDetails("@DOC1", "examt", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");
                FieldDetails("@DOC1", "exaltexp", "Allocated Expense", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, 20, "", false, "");

                ChildTableDM.Add("DOC1"); 
                CreateUserObject("XDOC", "Detailed Expenses", "XDOC", ChildTableDM, FindColumnDM, BoYesNoEnum.tYES, BoUDOObjType.boud_Document);

                FieldDetails("OQUT", "PRLNum", "PriceList Num", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDR", "PRLNum", "PriceList Num", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORRR", "PRLNum", "PriceList Num", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDN", "PRLNum", "PriceList Num", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPI", "PRLNum", "PriceList Num", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OINV", "PRLNum", "PriceList Num", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORIN", "PRLNum", "PriceList Num", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");

                
                // UDFS in All Marketing Documents
                FieldDetails("OQUT", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OQUT", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OQUT", "EXIMFD",  "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OQUT", "EXIMOC",  "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OQUT", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OQUT", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OQUT", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OQUT", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OQUT", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
 
                FieldDetails("OQUT", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OQUT", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OQUT", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OQUT", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OQUT", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OQUT", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OQUT", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OQUT", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OQUT", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OQUT", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OQUT", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OQUT", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                 

                FieldDetails("ODLN", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODLN", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODLN", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODLN", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODLN", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODLN", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODLN", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ODLN", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ODLN", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ODLN", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODLN", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODLN", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODLN", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODLN", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODLN", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODLN", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODLN", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODLN", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODLN", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODLN", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODLN", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                 

                FieldDetails("ORRR", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORRR", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORRR", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORRR", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORRR", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORRR", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORRR", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORRR", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORRR", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORRR", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORRR", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORRR", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORRR", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORRR", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORRR", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORRR", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORRR", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORRR", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORRR", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORRR", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORRR", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");


                FieldDetails("ORDN", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDN", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDN", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDN", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDN", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDN", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDN", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORDN", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORDN", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORDN", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORDN", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORDN", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORDN", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORDN", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORDN", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORDN", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORDN", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORDN", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORDN", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORDN", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORDN", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");


                FieldDetails("ODPI", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPI", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPI", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPI", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPI", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OQUT", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OQUT", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OQUT", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OQUT", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OQUT", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OQUT", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OQUT", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OQUT", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OQUT", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OQUT", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OQUT", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OQUT", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OQUT", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OQUT", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OQUT", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OQUT", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");


                FieldDetails("ORDR", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDR", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDR", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDR", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDR", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDR", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORDR", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORDR", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORDR", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORDR", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORDR", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORDR", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORDR", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORDR", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORDR", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORDR", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORDR", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORDR", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORDR", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORDR", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORDR", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");

                FieldDetails("OINV", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OINV", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OINV", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OINV", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OINV", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OINV", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OINV", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OINV", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OINV", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OINV", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OINV", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OINV", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OINV", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OINV", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OINV", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OINV", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OINV", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OINV", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OINV", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OINV", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OINV", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");


                FieldDetails("ORIN", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORIN", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORIN", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORIN", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORIN", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORIN", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORIN", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORIN", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORIN", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ODLN", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODLN", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODLN", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODLN", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODLN", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODLN", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODLN", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODLN", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODLN", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODLN", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODLN", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODLN", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");


                // UDFS in All Marketing Documents
                FieldDetails("OPQT", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPQT", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPQT", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPQT", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPQT", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPQT", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPQT", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPQT", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPQT", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPQT", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPQT", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPQT", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPQT", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPQT", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPQT", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPQT", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPQT", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPQT", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPQT", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPQT", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPQT", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPQT", "LCNO", "LC No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 50, "", false, "");

                FieldDetails("OPOR", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPOR", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPOR", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPOR", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPOR", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPOR", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPOR", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPOR", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPOR", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPOR", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPOR", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPOR", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPOR", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPOR", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPOR", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPOR", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPOR", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPOR", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPOR", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPOR", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPOR", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPOR", "LCNO", "LC No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 50, "", false, "");

                FieldDetails("OPDN", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPDN", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPDN", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPDN", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPDN", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPDN", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPDN", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPDN", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPDN", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPDN", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPDN", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPDN", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPDN", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPDN", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPDN", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPDN", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPDN", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPDN", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPDN", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPDN", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPDN", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPDN", "LCNO", "LC No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 50, "", false, "");

                FieldDetails("OPRR", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPRR", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPRR", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPRR", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPRR", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPRR", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPRR", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPRR", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPRR", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPRR", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPRR", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPRR", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPRR", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPRR", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPRR", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPRR", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPRR", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPRR", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPRR", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPRR", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPRR", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");

                FieldDetails("ORPD", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPD", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPD", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPD", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPD", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPD", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPD", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORPD", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORPD", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORPD", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORPD", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORPD", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORPD", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORPD", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORPD", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORPD", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORPD", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORPD", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORPD", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORPD", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORPD", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");

                FieldDetails("ODPO", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPO", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPO", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPO", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPO", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPO", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ODPO", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ODPO", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ODPO", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ODPO", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODPO", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODPO", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODPO", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODPO", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODPO", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODPO", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODPO", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODPO", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ODPO", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ODPO", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ODPO", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");

                FieldDetails("OPCH", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPCH", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPCH", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPCH", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPCH", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPCH", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("OPCH", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPCH", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPCH", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("OPCH", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPCH", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPCH", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPCH", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPCH", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPCH", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPCH", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPCH", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPCH", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPCH", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("OPCH", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OPCH", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("OPCH", "LCNO", "LC No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 50, "", false, "");

                FieldDetails("ORPC", "EXIMPOL", "Port of Loading", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPC", "EXIMPOD", "Port of Discharge", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPC", "EXIMFD", "Final Destination", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPC", "EXIMOC", "Origin Country", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPC", "EXIMFDC", "Final Destination Contry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPC", "EXIMPOR", "Port of Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("ORPC", "EXIMINCO", "Incoterm", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORPC", "EXIMPCGB", "Pre - Carriage By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORPC", "EXIMPCRB", "Pre - Carrier By", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 150, "", false, "");
                FieldDetails("ORPC", "EXIMBC", "Buyer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORPC", "EXIMBN", "Buyer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORPC", "EXIMBA", "Buyer Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORPC", "EXIMCC", "Consignee Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORPC", "EXIMCN", "Consignee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORPC", "EXIMCA", "Consignee Address", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORPC", "EXIMPNC1", "Notify PCode1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORPC", "EXIMNPN1", "Notify PName1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORPC", "EXIMNPA1", "Notify PAddress1", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");
                FieldDetails("ORPC", "EXIMNPC2", "Notify PCode2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", false, "");
                FieldDetails("ORPC", "EXIMNPN2", "Notify PName2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("ORPC", "EXIMNPA2", "Notify PAddress2", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Address, 250, "", false, "");

                FieldDetails("OITM", "EXIMLIC", "LIC Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("OITM", "EXIMLIN", "LIC Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");

                FieldDetails("OPRQ", "LCNO", "LC No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
               
                /* Tables for Price List */

                List<string> FindColumnPL = new List<string>();
                List<string> ChildTablePL = new List<string>();
                List<string> FindColumnGPL = new List<string>();
                List<string> ChildTableGPL = new List<string>();

                CreateTable("SMPL", "Price List", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                FieldDetails("@SMPL", "cardcode", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@SMPL", "cardname", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");
                FieldDetails("@SMPL", "freight", "Freight Charge", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@SMPL", "packing", "Packing Charge", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@SMPL", "factExp", "Factory Expense", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@SMPL", "addexp1", "Additional Charge 1", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@SMPL", "addexp2", "Additional Charge 2", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@SMPL", "addexp3", "Additional Charge 3", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@SMPL", "exchRate", "Exchange Rate", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 20, "", false, "");
                FieldDetails("@SMPL", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                FieldDetails("@SMPL", "docDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@SMPL", "disPer", "Discount %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 20, "", false, "");
                FieldDetails("@SMPL", "disAmt", "Discount Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@SMPL", "profPer", "Profit %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 20, "", false, "");
                FieldDetails("@SMPL", "profAmt", "Profit Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@SMPL", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");
                FieldDetails("@SMPL", "currency", "currency", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                
                CreateTable("MPL1", "Labmix Samples", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@MPL1", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@MPL1", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("@MPL1", "fgitemcode", "FG Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@MPL1", "fgitemname", "FG Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("@MPL1", "inoutno", "Inward/Outward No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");
                FieldDetails("@MPL1", "ioref", "Inward/Outward Ref No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");
                FieldDetails("@MPL1", "ptyref", "Party Ref No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");
                FieldDetails("@MPL1", "stdPrice", "Standard Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "untPrice", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "disPer", "Discount %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 20, "", false, "");
                FieldDetails("@MPL1", "disAmt", "Discount Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "profPer", "Profit %", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, 20, "", false, "");
                FieldDetails("@MPL1", "profAmt", "Profit Amount", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "freight", "Freight Charge", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "packing", "Packing Charge", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "factExp", "Factory Expense", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "addexp1", "Additional Charge 1", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "addexp2", "Additional Charge 2", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "addexp3", "Additional Charge 3", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "stdPriceLC", "Standard Price LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL1", "untPriceLC", "Unit Price LC", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                 
                CreateTable("MPL2", "Finished Goods", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@MPL2", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@MPL2", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");
                FieldDetails("@MPL2", "stdprice", "Standard Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL2", "updprice", "Updated Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@MPL2", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");
                
                ChildTablePL.Add("MPL1");
                ChildTablePL.Add("MPL2");
                FindColumnPL.Add("U_cardcode");
                FindColumnPL.Add("U_cardname");
                CreateUserObject("SMPL", "Price List", "SMPL", ChildTablePL, FindColumnPL, BoYesNoEnum.tNO, BoUDOObjType.boud_MasterData);

                CreateTable("FGPL", "FG Price List", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                FieldDetails("@FGPL", "title", "title", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 200, "", false, "");
                FieldDetails("@FGPL", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", false, "");
                FieldDetails("@FGPL", "docDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("GPL1", "FG Price Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                FieldDetails("@GPL1", "itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@GPL1", "itemname", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 250, "", false, "");
                FieldDetails("@GPL1", "price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                
                ChildTableGPL.Add("GPL1");
                CreateUserObject("FGPL", "FG Price List", "FGPL", ChildTableGPL, FindColumnGPL, BoYesNoEnum.tNO, BoUDOObjType.boud_MasterData);
                 
                // **** Email Automation Tables*****
                CreateTable("AUEM", "CS Email Auto", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                FieldDetails("@AUEM", "EmailFR", "Default From EmailID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@AUEM", "EmailCC", "Default CC EmailID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@AUEM", "EmailSUB", "Default Subject", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@AUEM", "EmailBDY", "Default Body", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@AUEM", "EmailRM", "Default Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@AUEM", "DepWise", "Department Wise", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@AUEM", "SeWise", "Sales Employee Wise", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@AUEM", "OwWise", "Owner Wise", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@AUEM", "FrmBPM", "From BP master", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@AUEM", "AAcnt", "All Active Contact", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@AUEM", "Defcnt", "Default contact person", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@AUEM", "cntTran", "Contact Person from Trans", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                 
                CreateTable("EMDW", "CS Email DocWise", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                FieldDetails("@EMDW", "DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EMDW", "EmailFR", "Default From EmailID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EMDW", "EmailCC", "Default CC EmailID", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@EMDW", "EmailSUB", "Default Subject", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EMDW", "EmailBDY", "Default Body", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EMDW", "EmailRM", "Default Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@EMDW", "DepWise", "Department Wise", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EMDW", "SeWise", "Sales Employee Wise", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EMDW", "OwWise", "Owner Wise", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EMDW", "FrmBPM", "From BP master", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EMDW", "AAcnt", "All Active Contact", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EMDW", "Defcnt", "Default contact person", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@EMDW", "cntTran", "Contact Person from Trans", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");

                /////////////////// 
                /// 

                // **** JOBWORK : Nature of Processing Master *****88
                CreateTable("JOPM", "JOPMaster", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                FieldDetails("@JOPM", "nopcode", "Nature of Processing Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JOPM", "nopname", "Nature of Processing Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JOPM", "attachment", "Attachment", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_Link, 250, "", false, "");
                FieldDetails("@JOPM", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@JOPM", "status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");

                List<string> FindColumnJOPM = new List<string>();
                List<string> ChildTableJOPM = new List<string>();
                string[] DatalistJOPM = { "U_nopcode", "U_nopname", "U_remarks", "U_status"};
                FindColumnJOPM.AddRange(DatalistJOPM);
                CreateUserObject("JOPM", "Nature of Processing Master", "JOPM", ChildTableJOPM, FindColumnJOPM, BoYesNoEnum.tNO, BoUDOObjType.boud_MasterData);
                 
                CreateTable("JOTR", "Jobwork Challan - Out", SAPbobsCOM.BoUTBTableType.bott_Document);
                FieldDetails("@JOTR", "CardCode", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JOTR", "CardName", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JOTR", "PoNum", "Purchase Order", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@JOTR", "PoDe", "PO DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@JOTR", "NumAtCard", "Vendor Ref No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JOTR", "WhsCode", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JOTR", "BPLId", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JOTR", "DocStatus", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@JOTR", "DocDate", "Doc Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@JOTR", "DocDueDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@JOTR", "nopcode", "Nature of Processing", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JOTR", "ApInvNum", "A/P Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JOTR", "ApInvDe", "A/P Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JOTR", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
               
                CreateTable("OTR1", "JCO-Finished Goods", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@OTR1", "PoNum", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR1", "ByProduct", "By Product", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR1", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR1", "Dscription", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("@OTR1", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR1", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@OTR1", "WhsCode", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR1", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@OTR1", "RecQty", "Received Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR1", "BalQty", "Balance Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR1", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                 
                CreateTable("OTR2", "JCO-Components", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@OTR2", "PoNum", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR2", "FGCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR2", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR2", "Dscription", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("@OTR2", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR2", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@OTR2", "FrmWhs", "From Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR2", "ToWhs", "To Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR2", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@OTR2", "IsueQty", "Issued Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity,20, "", false, "");
                FieldDetails("@OTR2", "BalQty", "Balance Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR2", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("OTR3", "JCO-linked Doc FG", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@OTR3", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR3", "BaseDocNo", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR3", "BaseDocEnt", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR3", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@OTR3", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("OTR4", "JCO-linked Doc Comp", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@OTR4", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR4", "BaseDocNo", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR4", "BaseDocEnt", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR4", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@OTR4", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                /*
                CreateTable("OTR5", "JCO-RFinished Goods", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@OTR5", "PoNum", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR5", "ByProduct", "By Product", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR5", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR5", "Dscription", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("@OTR5", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR5", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@OTR5", "WhsCode", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR5", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@OTR5", "RecQty", "Received Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR5", "BalQty", "Balance Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false,"");
                FieldDetails("@OTR5", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("OTR6", "JCO-RComponents", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@OTR6", "PoNum", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR6", "FGCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR6", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR6", "Dscription", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("@OTR6", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR6", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@OTR6", "FrmWhs", "From Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR6", "ToWhs", "To Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR6", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@OTR6", "IsueQty", "Issued Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR6", "BalQty", "Balance Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@OTR6", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("OTR7", "JCO-Rlinked Doc FG", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@OTR7", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR7", "BaseDocNo", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR7", "BaseDocEnt", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR7", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@OTR7", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("OTR8", "JCO-Rlinked Doc Comp", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@OTR8", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@OTR8", "BaseDocNo", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR8", "BaseDocEnt", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@OTR8", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@OTR8", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                */
                List<string> FindColumnJOTR = new List<string>();
                List<string> ChildTableJOTR = new List<string>(); 
                string[] DatalistChildTableJOTR = { "OTR1", "OTR2", "OTR3", "OTR4" };
                ChildTableJOTR.AddRange(DatalistChildTableJOTR);
                string[] DatalistJOTR = { "U_CardCode", "U_CardName", "U_PoNum", "U_NumAtCard", "U_WhsCode", "U_ApInvNum" };
                FindColumnJOTR.AddRange(DatalistJOTR);
                CreateUserObject("JOTR", "Jobwork Challan - Out", "JOTR", ChildTableJOTR, FindColumnJOTR, BoYesNoEnum.tYES, BoUDOObjType.boud_Document);

                CreateTable("JITR", "Jobwork Challan - In", SAPbobsCOM.BoUTBTableType.bott_Document);
                FieldDetails("@JITR", "CardCode", "BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JITR", "CardName", "BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JITR", "SoNum", "Sales Order", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@JITR", "SoDe", "SO DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@JITR", "NumAtCard", "Vendor Ref No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JITR", "WhsCode", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JITR", "BPLId", "Branch", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@JITR", "DocStatus", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", false, "");
                FieldDetails("@JITR", "DocDate", "Doc Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@JITR", "DocDueDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@JITR", "nopcode", "Nature of Processing", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JITR", "ArInvNum", "A/r Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JITR", "ArInvDe", "A/r Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", false, "");
                FieldDetails("@JITR", "remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                 
                CreateTable("ITR1", "JCO-Finished Goods", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@ITR1", "PoNum", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@ITR1", "ByProduct", "By Product", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@ITR1", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@ITR1", "Dscription", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("@ITR1", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@ITR1", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@ITR1", "WhsCode", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@ITR1", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@ITR1", "RecQty", "Received Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@ITR1", "BalQty", "Balance Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@ITR1", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("ITR2", "JCO-Components", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@ITR2", "PoNum", "PO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@ITR2", "FGCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@ITR2", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@ITR2", "Dscription", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("@ITR2", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@ITR2", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("@ITR2", "FrmWhs", "From Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@ITR2", "ToWhs", "To Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@ITR2", "Price", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");
                FieldDetails("@ITR2", "IsueQty", "Issued Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@ITR2", "BalQty", "Balance Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("@ITR2", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("ITR3", "JCO-linked Doc FG", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@ITR3", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@ITR3", "BaseDocNo", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@ITR3", "BaseDocEnt", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@ITR3", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@ITR3", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                CreateTable("ITR4", "JCO-linked Doc Comp", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                FieldDetails("@ITR4", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@ITR4", "BaseDocNo", "DocNum", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@ITR4", "BaseDocEnt", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100, "", false, "");
                FieldDetails("@ITR4", "DocDate", "DocDate", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");
                FieldDetails("@ITR4", "remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 0, "", false, "");

                //FieldDetailsLink("POR1", "ItemCode_CSJW", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("POR1", "ItemCode_CSJW", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("POR1", "Desc_CSJW", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("POR1", "Qty_CSJW", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("POR1", "UOM_CSJW", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("POR1", "Price_CSJW", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");

                //FieldDetailsLink("RDR1", "ItemCode_CSJW", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("RDR1", "ItemCode_CSJW", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("RDR1", "Desc_CSJW", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 155, "", false, "");
                FieldDetails("RDR1", "Qty_CSJW", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, 20, "", false, "");
                FieldDetails("RDR1", "UOM_CSJW", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 25, "", false, "");
                FieldDetails("RDR1", "Price_CSJW", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 20, "", false, "");

                FieldDetailsLink("OCRD", "DefWhs", "Jobwork Whs", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, ""); 
                FieldDetails("OWTR", "JWODe", "Jobwork Out No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("OIGE", "JWODe", "Jobwork Out No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("OIGN", "JWODe", "Jobwork Out No", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                
                List<string> FindColumnJITR = new List<string>();
                List<string> ChildTableJITR = new List<string>();
                string[] DatalistChildTableJITR = { "ITR1", "ITR2", "ITR3", "ITR4" };
                ChildTableJITR.AddRange(DatalistChildTableJITR);
                string[] DatalistJITR = { "U_CardCode", "U_CardName", "U_SoNum", "U_NumAtCard", "U_WhsCode", "U_ArInvNum" };
                FindColumnJITR.AddRange(DatalistJITR);
                CreateUserObject("JITR", "Jobwork Challan - In", "JITR", ChildTableJITR, FindColumnJITR, BoYesNoEnum.tYES, BoUDOObjType.boud_Document);
                 
                CreateTable("JOREL", "Jobwork Relations", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                FieldDetails("@JOREL", "JWOId", "Jobwork Out Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@JOREL", "JWIId", "Jobwork IN Id", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@JOREL", "GI", "Goods Issue", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");
                FieldDetails("@JOREL", "GR", "Goods Receipt", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", false, "");


                

                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message.ToString() + ":    : Add-On database tables creation fail!");
                return false;
            }
        }

        private bool CreateUserObject(string CodeID, string Name, string TableName, List<string> ChildTable, List<string> FindColumn, SAPbobsCOM.BoYesNoEnum ManageSeries, SAPbobsCOM.BoUDOObjType Type) //used for registration of user defined table
        {
            try
            {
                int lRetCode = 0;
                string sErrMsg = null;
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (oUserObjectMD == null)
                {
                    oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));
                }

                if (oUserObjectMD.GetByKey(CodeID) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                    oUserObjectMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                }
                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.TableName = TableName;
                oUserObjectMD.ObjectType = Type;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

                if (Type == SAPbobsCOM.BoUDOObjType.boud_MasterData)
                {
                    //oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObjectMD.ManageSeries = ManageSeries;
                }
                else if (Type == SAPbobsCOM.BoUDOObjType.boud_Document)
                {
                    oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObjectMD.ManageSeries = ManageSeries;

                }
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;

                foreach (string ChildTableName in ChildTable)
                {
                    oUserObjectMD.ChildTables.TableName = ChildTableName;
                    oUserObjectMD.ChildTables.Add();
                }

                // int i=1;
                foreach (string Col in FindColumn)
                {
                    //oUserObjectMD.FindColumns.SetCurrentLine(i);
                    oUserObjectMD.FindColumns.ColumnAlias = Col;
                    oUserObjectMD.FindColumns.Add();
                } 

                lRetCode = oUserObjectMD.Add();
                if (lRetCode != 0)
                {
                    if (lRetCode == -1)
                    { }
                    else
                    {
                        SBOMain.oCompany.GetLastError(out lRetCode, out sErrMsg);
                        SBOMain.SBO_Application.StatusBar.SetText("Object:" + CodeID + "Creation Error: " + sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                else
                {
                    SBOMain.SBO_Application.StatusBar.SetText("Object:" + CodeID + "Created Successfully..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                    oUserObjectMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                return true;
            }
            catch
            {
                return false;
            }
        } 

        private bool CreateTable(string TableName, string TableDesc, SAPbobsCOM.BoUTBTableType TableType)
        {
            try
            {
                int errCode;
                string ErrMsg = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (oUserTablesMD == null)
                {
                    oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
                }
                if (oUserTablesMD.GetByKey(TableName) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    return true;
                } 
                oUserTablesMD.TableName = TableName;
                oUserTablesMD.TableDescription = TableDesc;
                oUserTablesMD.TableType = TableType;

                long err = oUserTablesMD.Add();
                if (err != 0)
                {
                    SBOMain.oCompany.GetLastError(out errCode, out ErrMsg);
                }
                if (err == 0)
                {
                    SBOMain.SBO_Application.StatusBar.SetText("Table Created : " + TableDesc + "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool FieldDetails(string TableName, string FieldName, string FieldDesc, SAPbobsCOM.BoFieldTypes FieldType, SAPbobsCOM.BoFldSubTypes FieldSubType, int FieldSize, string ValidValues, bool Mandatory, string DefaultVal)
        {
            if (FieldExist(TableName, FieldName) == false)
            {
                string ErrMsg;
                int errCode;
                int IRetCode;
                oUserFieldsMD = null; 
                try
                {
                    GC.Collect();
                    oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
                    oUserFieldsMD.TableName = TableName;
                    oUserFieldsMD.Name = FieldName;
                    oUserFieldsMD.Description = FieldDesc;
                     
                    oUserFieldsMD.Type = FieldType;
                    oUserFieldsMD.SubType = FieldSubType;
                    oUserFieldsMD.EditSize = FieldSize; 
                    
                    /*if(TableName == "OQUT" || TableName == "ORDR")
                    {
                        oUserFieldsMD.LinkedTable = "SMPL";
                    }*/
                        if (ValidValues != "")
                    {
                        switch (ValidValues)
                        {
                            case "MTYPE":
                                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                              
                                oUserFieldsMD.EditSize = FieldSize;

                                oUserFieldsMD.ValidValues.Value = "BOPP";
                                oUserFieldsMD.ValidValues.Description = "BOPP";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "PVC";
                                oUserFieldsMD.ValidValues.Description = "PVC";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "PET";
                                oUserFieldsMD.ValidValues.Description = "PET";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "HST";
                                oUserFieldsMD.ValidValues.Description = "HST";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "";
                                break;

                            case "AP":
                                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                              
                                oUserFieldsMD.EditSize = FieldSize;

                                oUserFieldsMD.ValidValues.Value = "Pending";
                                oUserFieldsMD.ValidValues.Description = "Pending";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "Approve";
                                oUserFieldsMD.ValidValues.Description = "Approve";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "Pending";
                                break;
                        }
                    }

                    if (Mandatory)
                    {
                        oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    if (DefaultVal != "")
                    {
                        oUserFieldsMD.DefaultValue = DefaultVal;
                    }

                    // Add the field to the table
                    IRetCode = oUserFieldsMD.Add();
                    if (IRetCode != 0)
                    {
                        if (IRetCode == -2035 || IRetCode == -1120)
                        {
                            return false;
                        }
                        else
                        {
                            SBOMain.oCompany.GetLastError(out errCode, out ErrMsg);
                            SBOMain.SBO_Application.SetStatusBarMessage("Error : " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                    else
                    {
                        SBOMain.SBO_Application.SetStatusBarMessage("Field Created in : " + TableName + " As : " + FieldDesc, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    oUserFieldsMD = null;
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
                finally
                {
                    ErrMsg = null; errCode = 0; IRetCode = 0;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }
            return true;
        }

        private bool FieldDetailsLink(string TableName, string FieldName, string FieldDesc, SAPbobsCOM.BoFieldTypes FieldType, SAPbobsCOM.BoFldSubTypes FieldSubType, int FieldSize, string ValidValues, bool Mandatory, string DefaultVal)
        {
            if (FieldExist(TableName, FieldName) == false)
            {
                string ErrMsg;
                int errCode;
                int IRetCode;
                oUserFieldsMD = null;
                try
                {
                    GC.Collect();
                    oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
                    oUserFieldsMD.TableName = TableName;
                    oUserFieldsMD.Name = FieldName;
                    oUserFieldsMD.Description = FieldDesc;
                      
                    oUserFieldsMD.Type = FieldType;
                    oUserFieldsMD.SubType = FieldSubType;
                    oUserFieldsMD.EditSize = FieldSize;

                    if(FieldDesc == "Item Code") 
                    { 
                        oUserFieldsMD.LinkedSystemObject = UDFLinkedSystemObjectTypesEnum.ulItems;
                    }
                    else if(FieldDesc == "Jobwork Whs")
                    {
                        oUserFieldsMD.LinkedSystemObject = UDFLinkedSystemObjectTypesEnum.ulWarehouses;
                    }

                    //oUserFieldsMD.LinkedTable = "SMPL";

                    /*if(TableName == "OQUT" || TableName == "ORDR")
                    {
                        oUserFieldsMD.LinkedTable = "SMPL";
                    }*/
                    if (ValidValues != "")
                    {
                        switch (ValidValues)
                        {
                            case "MTYPE":
                                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                              
                                oUserFieldsMD.EditSize = FieldSize;

                                oUserFieldsMD.ValidValues.Value = "BOPP";
                                oUserFieldsMD.ValidValues.Description = "BOPP";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "PVC";
                                oUserFieldsMD.ValidValues.Description = "PVC";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "PET";
                                oUserFieldsMD.ValidValues.Description = "PET";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "HST";
                                oUserFieldsMD.ValidValues.Description = "HST";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "";
                                break;

                            case "AP":
                                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha; // alphanumeric type                              
                                oUserFieldsMD.EditSize = FieldSize;

                                oUserFieldsMD.ValidValues.Value = "Pending";
                                oUserFieldsMD.ValidValues.Description = "Pending";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.ValidValues.Value = "Approve";
                                oUserFieldsMD.ValidValues.Description = "Approve";
                                oUserFieldsMD.ValidValues.Add();

                                oUserFieldsMD.DefaultValue = "Pending";
                                break;
                        }
                    }

                    if (Mandatory)
                    {
                        oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    if (DefaultVal != "")
                    {
                        oUserFieldsMD.DefaultValue = DefaultVal;
                    }

                    // Add the field to the table
                    IRetCode = oUserFieldsMD.Add();
                    if (IRetCode != 0)
                    {
                        if (IRetCode == -2035 || IRetCode == -1120)
                        {
                            return false;
                        }
                        else
                        {
                            SBOMain.oCompany.GetLastError(out errCode, out ErrMsg);
                            SBOMain.SBO_Application.SetStatusBarMessage("Error : " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                    else
                    {
                        SBOMain.SBO_Application.SetStatusBarMessage("Field Created in : " + TableName + " As : " + FieldDesc, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    oUserFieldsMD = null;
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
                finally
                {
                    ErrMsg = null; errCode = 0; IRetCode = 0;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }
            return true;
        }

        private bool FieldExist(string TableName, string ColumnName)
        {
            SAPbobsCOM.Recordset oRecordSet = default(SAPbobsCOM.Recordset);
            oRecordSet = SBOMain.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + TableName + "' AND \"AliasID\" = '" + ColumnName + "'");
                if ((Convert.ToInt32(oRecordSet.Fields.Item(0).Value) == 0))
                {
                    oRecordset = null;
                    return false;
                }
                else
                {
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
                    oRecordset = null;
                    return true;
                }
                //Return (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0)
            }
            catch (Exception ex)
            {
                throw ex;
            }
            //finally
            //{
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            //    GC.Collect();
            //}
        }
         
    }
}
