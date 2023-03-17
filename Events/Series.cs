using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSuteConnect.Events
{
    public static class Series
    {
        public static void SeriesCombo(string OBJ, string ColId)
        {
            try
            {
                SAPbouiCOM.Form oForm; 
                SAPbouiCOM.ComboBox oCombo;
                oCombo = null;
                oForm = SBOMain.SBO_Application.Forms.ActiveForm;
                oCombo = oForm.Items.Item(ColId).Specific;

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    oCombo.ValidValues.LoadSeries(OBJ, SAPbouiCOM.BoSeriesMode.sf_Add);
                else
                    oCombo.ValidValues.LoadSeries(OBJ, SAPbouiCOM.BoSeriesMode.sf_View);

                oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item(ColId).DisplayDesc = true;
            }
            catch (Exception ex)
            {
                SBOMain.SBO_Application.StatusBar.SetText("Series : " + ex, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
