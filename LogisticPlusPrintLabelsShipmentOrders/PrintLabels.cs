using Mantis.LVision.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LogisticPlusPrintLabelsShipmentOrders
{
    public class PrintLabels : IReportInfo_Version_1_1
    {
        public object SelectReportInfo(object frm, int ReportType, object ConstructParam, DataSet dsReportSource)
        {
            var lvBasicForm = (LVBasicForm)frm;
            new PrintForm(lvBasicForm, m_SelectedItemsIDs, "exec LP_REP_CheckList_OrderShipmentID_Billa_proc ").ShowDialog();
            return true;
        }

        public int[] m_SelectedItemsIDs;


        int[] IReportInfo_Version_1_1.SelectedItemIDs
        {
            get
            {
                return this.m_SelectedItemsIDs;
            }
            set
            {
                this.m_SelectedItemsIDs = value;
            }
        }

        public int[] SelectedItemIDs
        {
            get
            {
                return this.m_SelectedItemsIDs;
            }
            set
            {
                this.m_SelectedItemsIDs = value;
            }
        }
    }
}
