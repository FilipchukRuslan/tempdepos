using System;
using System.Data;
using System.Data.SqlClient;

namespace LogisticPlusPrintLabelsShipmentOrders
{
    public static class Logger
    {
        public static void InsertLogIntoDB(string ConnectionString, int orderID, string ordercode, string user, int qty, int depositorID)
        {
            using (var sqlConn = new SqlConnection(ConnectionString))
            {
                var cmd = sqlConn.CreateCommand();
                cmd.CommandText = $@"INSERT INTO LP_LogPrintLabelOrders (lplo_ordID, lplo_ordCode, lplo_printDate, lplo_stickerQty, lplo_userName, lplo_depID) 
values (@orderID, @ordercode, @date, @qty, @user, @depositorID)";
                cmd.Parameters.Add(new SqlParameter("@orderID", orderID));
                cmd.Parameters.Add(new SqlParameter("@ordercode", ordercode));
                cmd.Parameters.Add(new SqlParameter("@date", DateTime.Now));
                cmd.Parameters.Add(new SqlParameter("@qty", qty));
                cmd.Parameters.Add(new SqlParameter("@user", user));
                cmd.Parameters.Add(new SqlParameter("@depositorID", depositorID));

                sqlConn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        public static void UpdateLogInDB(string ConnectionString, string ordercode, string user, int qty)
        {
            using (var sqlConn = new SqlConnection(ConnectionString))
            {
                var cmd = sqlConn.CreateCommand();
                cmd.CommandText = $@"UPDATE LP_LogPrintLabelOrders 
SET 
lplo_stickerQty = @qty, 
lplo_userName = @user, 
lplo_printDate = @date
WHERE 
lplo_ordCode = @ordercode";
                cmd.Parameters.Add(new SqlParameter("@ordercode", ordercode));
                cmd.Parameters.Add(new SqlParameter("@user", user));
                cmd.Parameters.Add(new SqlParameter("@qty", qty));
                cmd.Parameters.Add(new SqlParameter("@date", DateTime.Now));
                
                sqlConn.Open();
                cmd.ExecuteNonQuery();
            }
        }

        public static bool CheckLog(string ConnectionString, string ordercode)
        {
            try
            {
                using (var sqlConn = new SqlConnection(ConnectionString))
                {
                    var cmd = sqlConn.CreateCommand();
                    cmd.CommandText = $"SELECT lplo_ID FROM LP_LogPrintLabelOrders With (Nolock) WHERE lplo_ordCode = @lplo_ordCode";
                    cmd.Parameters.Add(new SqlParameter("@lplo_ordCode", ordercode));
                    sqlConn.Open();
                    var result = cmd.ExecuteScalar();
                    return result != null && result != DBNull.Value;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}
