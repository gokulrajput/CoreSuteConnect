using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreSuteConnect.Events;
using SAPbouiCOM;
using SAPbobsCOM;
using CoreSuteConnect.Class.Common; 
using System.Xml;
using System.Net.Http.Headers; 

namespace CoreSuteConnect.Class.AUTOEMAIL
{
    public class SBOMainAUTMAIL
    {
        clsEmailAutomation clsEmailAutomation = new clsEmailAutomation();  

        public bool SubItemEvent(ref ItemEvent pVal, ref bool BubbleEvent, string FormUID)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.FormTypeEx)
                { 
                    case "frmEmailAutomation":
                        clsEmailAutomation.ItemEvent(ref pVal, ref BubbleEvent, pVal.FormUID);
                        break; 
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(ex.Message.ToString(), 1, "ok", "", "");
            }
            return BubbleEvent;
        }

    }
}
