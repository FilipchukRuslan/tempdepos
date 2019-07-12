using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using Mantis.LVision.LabelUtils;
using Mantis.LVision.Win32;


namespace LogisticPlusPrintLabelsShipmentOrders
{
    public class PrintLabels : IReportInfo_Version_1_1
    {
        public int[] m_SelectedItemsIDs;
        public int Number { get; set; }
        Form printForm;
        Button buttonOK;
        Button buttonCancel;
        NumericUpDown numericUpDown1;
        int GetSumNumbers(SqlConnection sqlConnection)
        {
            int temp = 0;
            foreach (var mSelectedItemsId in m_SelectedItemsIDs)
            {
                var command = sqlConnection.CreateCommand();
                command.CommandText = string.Format(@"SELECT t2.MAX_COUNT AS NUM
            FROM (SELECT vlsc.shc_ID AS ID1, vlsc2.shc_ID AS ID2, ROW_NUMBER() OVER (PARTITION BY vlsc.shc_ID
            ORDER BY vlsc2.shc_ID) AS ROW_NUM,
            (SELECT COUNT(vlsc3.shc_ID) AS cnt
            FROM LV_OrderShipment vlos3 
			LEFT JOIN LV_ShipContainer vlsc3 ON vlsc3.shc_OrderShipmentID = vlos3.ost_ID
            WHERE vlos3.ost_OrderID = vlo.ord_ID) MAX_COUNT
            FROM LV_ShipContainer vlsc 
			LEFT JOIN LV_OrderShipment vlos ON vlos.ost_ID = vlsc.shc_OrderShipmentID 
			LEFT JOIN LV_Order vlo ON vlo.ord_ID = vlos.ost_OrderID 
			LEFT JOIN LV_OrderShipment vlos2 ON vlos2.ost_OrderID = vlo.ord_ID 
			LEFT JOIN LV_ShipContainer vlsc2 ON vlsc2.shc_OrderShipmentID = vlos2.ost_ID
            WHERE vlsc.shc_ID = {0}) t2
WHERE t2.ID1 = t2.ID2", mSelectedItemsId);
                var obj = command.ExecuteScalar();
                temp += (int)obj;
            }
            return temp;
        }

        private void DrawForm(Form form, int number)
        {
            buttonOK = new Button();
            buttonCancel = new Button();
            numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            form.SuspendLayout();

            buttonOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            buttonOK.Location = new System.Drawing.Point(89, 183);
            buttonOK.Name = "buttonOK";
            buttonOK.Size = new System.Drawing.Size(113, 40);
            buttonOK.TabIndex = 1;
            buttonOK.Text = "Печать";
            buttonOK.UseVisualStyleBackColor = true;
            buttonOK.Click += new System.EventHandler(buttonOK_Click);
            // 
            // numericUpDown1
            // 
            numericUpDown1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            numericUpDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel, ((byte)(204)));
            numericUpDown1.Location = new System.Drawing.Point(50, 71);
            numericUpDown1.Margin = new System.Windows.Forms.Padding(15);
            numericUpDown1.Name = "numericUpDown1";
            numericUpDown1.Size = new System.Drawing.Size(198, 28);
            numericUpDown1.TabIndex = 2;
            numericUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            numericUpDown1.Value = number;
            numericUpDown1.ValueChanged += new System.EventHandler(numericUpDown1_ValueChanged);
            // 
            // PrintForm
            // 
            form.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            form.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            form.BackColor = System.Drawing.SystemColors.Menu;
            form.ClientSize = new System.Drawing.Size(284, 261);
            form.Controls.Add(this.numericUpDown1);
            form.Controls.Add(this.buttonOK);
            form.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel, ((byte)(204)));
            form.MinimumSize = new System.Drawing.Size(300, 260);
            form.Name = "PrintForm";
            form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            form.Text = "Введите количество";
            form.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            form.ResumeLayout(false);

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Number = (int)(sender as NumericUpDown).Value;
        }
        

        private void buttonOK_Click(object sender, EventArgs e)
        {
            Number = Convert.ToInt32(numericUpDown1.Value);
        }

        public object SelectReportInfo(object frm, int ReportType, object ConstructParam, DataSet dsReportSource)
        {
            try
            {
                
                var p = new PrinterSettings
                {
                    //"Xerox WC5765"; ZDesigner 105SL 203DPI
                    PrinterName = "ZDesigner 105SL 203DPI"
                };
                var con = ((LVBasicForm)frm).ConnectionString;

                using (var connection = new SqlConnection(con))
                {
                    connection.Open();
                    printForm = new Form();
                    DrawForm(printForm, GetSumNumbers(connection));
                    var res = printForm.ShowDialog();
                    if (res == DialogResult.Cancel)
                    {
                        return null;
                    }
                    else
                    {
                        printForm.Close();
                    }
                    for (int i = 0; i < m_SelectedItemsIDs.Length; i++)
                    {
                        #region query
                        var stringQuery = string.Format(

@"SELECT 
lsc.shc_ID, 
lo.ord_Code AS OrderCode, 
lo.ord_ID AS OrderID, 
ISNULL(lsc2.stc_SSCC, lsc.shc_SSCC) AS SSCC, 
ls.shp_Code AS ShipmentCode, 
lc2.cmp_ShortName AS CompanyShortName, 
lc2.cmp_Address AS CompanyAddress, 
lo.ord_InputDate AS OrderInputDate, 
t .UserName, t2.NUM as Qty, 
ld.dep_Code AS DepCode, 
ld.dep_ID AS DepID, 
lc3.cmp_ShortName AS DepShortName, 
lo.ord_Memo AS OrdMemo

FROM LV_ShipContainer lsc 
LEFT JOIN LV_StockContainer lsc2 ON lsc2.stc_ID = lsc.shc_ContainerID 
LEFT JOIN LV_OrderShipment los ON los.ost_ID = lsc.shc_OrderShipmentID
LEFT JOIN  LV_Order lo ON lo.ord_ID = los.ost_OrderID 
LEFT JOIN LV_Customer lc ON lc.cus_ID = lo.ord_CustomerID 
LEFT JOIN  LV_Company lc2 ON lc2.cmp_ID = lc.cus_CompanyID 
LEFT JOIN  LV_Shipment ls ON ls.shp_ID = los.ost_ShipmentID 
LEFT JOIN LV_Depositor ld ON ld.dep_ID = lo.ord_DepositorID 
LEFT JOIN  LV_Company lc3 ON lc3.cmp_ID = ld.dep_CompanyID
		 CROSS APPLY
            (SELECT CAST(t2.ROW_NUM AS NVARCHAR) + '/' + CAST({1} AS NVARCHAR) AS NUM --CAST(t2.MAX_COUNT AS NVARCHAR) AS NUM
            FROM (SELECT vlsc.shc_ID AS ID1, vlsc2.shc_ID AS ID2, ROW_NUMBER() OVER (PARTITION BY vlsc.shc_ID
            ORDER BY vlsc2.shc_ID) AS ROW_NUM,
            (SELECT COUNT(vlsc3.shc_ID) AS cnt
            FROM LV_OrderShipment vlos3 
			LEFT JOIN LV_ShipContainer vlsc3 ON vlsc3.shc_OrderShipmentID = vlos3.ost_ID
            WHERE vlos3.ost_OrderID = vlo.ord_ID) MAX_COUNT
            FROM LV_ShipContainer vlsc 
			LEFT JOIN LV_OrderShipment vlos ON vlos.ost_ID = vlsc.shc_OrderShipmentID 
			LEFT JOIN LV_Order vlo ON vlo.ord_ID = vlos.ost_OrderID 
			LEFT JOIN LV_OrderShipment vlos2 ON vlos2.ost_OrderID = vlo.ord_ID 
			LEFT JOIN LV_ShipContainer vlsc2 ON vlsc2.shc_OrderShipmentID = vlos2.ost_ID
            WHERE vlsc.shc_ID = lsc.shc_ID) t2
WHERE t2.ID1 = t2.ID2) t2 

CROSS APPLY
(SELECT cp.per_Code + ' - ' + cp.per_LastName + ' ' + cp.per_FirstName AS UserName, lh.hst_Name AS HostName
	FROM LV_Session ls 
	LEFT JOIN LV_Host lh ON lh.Hst_ID = ls.ses_HostID 
	LEFT JOIN LV_Users lu ON lu.usr_ID = ls.ses_UserID 
	LEFT JOIN COM_Person cp ON cp.per_ID = lu.usr_PersonID
	WHERE ls.ses_StartTime =
                (SELECT MAX(ls.ses_StartTime)
                    FROM LV_Session ls 
					LEFT JOIN LV_Host lh ON lh.Hst_ID = ls.ses_HostID
                    WHERE lh.hst_Name = (SELECT [host_name]
                    FROM sys.dm_exec_sessions
                    WHERE session_id = @@SPID)) AND lh.hst_Name =
                (SELECT [host_name]
                    FROM sys.dm_exec_sessions
                    WHERE session_id = @@SPID)) t
WHERE lo.ord_Code IN (SELECT OrderCode
FROM 
  LP_PrintShipContainerLabel_view 
WHERE 
  shc_ID = {0})", m_SelectedItemsIDs[i], Number);
                        #endregion

                        var adapter = new SqlDataAdapter(stringQuery, connection);
                        var ds = new DataSet();

                        adapter.Fill(ds);

                        var ordCode = ds.Tables[0].Rows[0]["OrderCode"].ToString();
                        var user = ((LVBasicForm)frm).UserName;
                        var depID = Convert.ToInt32(ds.Tables[0].Rows[0]["DepID"]);
                        var ordID = Convert.ToInt32(ds.Tables[0].Rows[0]["OrderID"]);

                        if (Logger.CheckLog(con, ordCode))
                        {
                            Logger.UpdateLogInDB(con, ordCode, user, Number);
                        }
                        else
                        {
                            Logger.InsertLogIntoDB(con, ordID, ordCode, user, Number, depID);
                        }

                        for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                        {
                            if (ds.Tables[0].Rows.Count < Number)
                            {
                                //добавление упаковочного места(SSCC)
                                var newR = ds.Tables[0].NewRow();
                                newR.ItemArray = ds.Tables[0].Rows[n].ItemArray;
                                newR["Qty"] = $"{ds.Tables[0].Rows.Count + 1}/{Number}";
                                ds.Tables[0].Rows.InsertAt(newR, ds.Tables[0].Rows.Count);
                            }
                            #region printLabel
                            var resString = string.Format(
@"
^XA
^ASN,20,20^FO20,8^FDУпаковочная этикетка отгрузки:^FS
^ASN,30,30^FO31,45^FDРасходный заказ: {0}^FS
^BY3^FO41,88^BCN,50,Y,N,N^FD>: {1}^FS
^BY3^FO41,173^BCN,50,Y,N,N^FD>: {2}^FS
^ASN,30,30^FO20,257^FDВладелец товара: {3} - {4}^FS
^FO8,295^GB782,3,3,B,0^FS
^ASN,33,22^FO20,303^FDПолучатель: {5}^FS
^ASN,33,22^FO20,346^FDАдрес: {6}^FS
^ASN,22,22^FO20,388^FDПользователь: {7}^FS
^ASN,22,22^FO450,388^FDДата ввода: {8}^FS
^ASN,33,22^FO20,436^FB650,1,,C,^FDПаллета: {9}^FS
^XZ
",
                            ds.Tables[0].Rows[n]["OrderCode"],
                            ds.Tables[0].Rows[n]["SSCC"],
                            ds.Tables[0].Rows[n]["ShipmentCode"],
                            ds.Tables[0].Rows[n]["DepCode"],
                            ds.Tables[0].Rows[n]["DepShortName"],
                            ds.Tables[0].Rows[n]["CompanyShortName"],
                            ds.Tables[0].Rows[n]["CompanyAddress"],
                            ((LVBasicForm)frm).UserName,
                            DateTime.Now,
                            ds.Tables[0].Rows[n]["Qty"]
                            )
                            ;
                            #endregion

                            stringQuery = String.Format("UPDATE LV_PrintLabel SET prl_Content = '{0}' WHERE prl_ID = 302", resString);
                            var command = connection.CreateCommand();
                            command = connection.CreateCommand();
                            command.CommandText = stringQuery;
                            command.ExecuteNonQuery();

                            /*Запускаем на печать*/
                            var hashtable = new Hashtable();
                            var infoArray = new PrintCrystalInfo[10];
                            int numOfCopies = 1;

                            PrintSettings.PrintLabelContainer(((LVBasicForm)frm), 302, 1, 1.0, ref numOfCopies, ref hashtable, p, ref infoArray);
                        }
                    }

                    connection.Close();
                    MessageBox.Show("Выполнено", "Выполнено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        int[] IReportInfo_Version_1_1.SelectedItemIDs
        {
            get
            {
                return m_SelectedItemsIDs;
            }
            set
            {
                m_SelectedItemsIDs = value;
            }
        }

        public int[] SelectedItemIDs
        {
            get
            {
                return m_SelectedItemsIDs;
            }
            set
            {
                m_SelectedItemsIDs = value;
            }
        }
    }
}
