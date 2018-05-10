using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace WinJob
{
    class Program
    {
        static string FilePath = ConfigurationManager.AppSettings["FilePath"];

        string ftpip = ConfigurationManager.AppSettings["ftpip"];
        string ftpusername = ConfigurationManager.AppSettings["ftpusername"];
        string ftppassword = ConfigurationManager.AppSettings["ftppassword"];

        SQLHelper SQLHelper = new SQLHelper();

        

        static void Main(string[] args)
        {
            if (ConfigurationManager.AppSettings["AutoRun"].ToString().Trim() == "Y")
            {
                Program p = new Program();
                p.getPOInfor();
            }
        }

        #region Pur_Po_ListMatExport

        private void getPOInfor()
        {
            if (!Directory.Exists(FilePath))
            {
                Directory.CreateDirectory(FilePath);
            }
            SFTPHelper sftp = new SFTPHelper(ftpip, ftpusername, ftppassword);

            DataSet ds = new DataSet();
            SqlParameter[] param = new SqlParameter[]
             {
                   new SqlParameter("@flag","PO")
             };
            ds = SQLHelper.GetDataSet("usp_Winjob", param);
            DataTable dt = ds.Tables[0];
            DataTable dt_dtl= ds.Tables[1];

            string filename = "";
            foreach (DataRow dr in dt.Rows)
            {
                filename = dr["PoDomain"].ToString() + "PUR" + DateTime.Now.ToString("yyyyMMddHHmmssfff");
                DataRow[] datarows = dt_dtl.Select("PONo='" + dr["PONo"].ToString() + "'");
                Pur_Po_ListMatExport(datarows, filename, dr["PONo"].ToString(), sftp);
            }
        }

        private int Pur_Po_ListMatExport(DataRow[] drs, string filename,string pono, SFTPHelper sftp)
        {
            var result = 0;
            try
            {
                FileStream fs = new FileStream(FilePath + @"\" + filename + ".txt", FileMode.Append, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding("UTF-8"));

                for (int i = 0; i < drs.Length; i++)
                {
                    sw.Write(drs[i]["PoDomain"].ToString()); sw.Write(",");//域
                    sw.Write(drs[i]["pur"].ToString()); sw.Write(",");//PUR固定
                    sw.Write(drs[i]["PONo"].ToString()); sw.Write(",");//采购订单号
                    sw.Write(drs[i]["PoVendorId"].ToString()); sw.Write(",");//采购供应商代码
                    sw.Write(drs[i]["wlh"].ToString()); sw.Write(",");//请购物料号
                    sw.Write(drs[i]["PurQty"].ToString()); sw.Write(",");//数量
                    sw.Write(drs[i]["NoTaxPrice"].ToString()); sw.Write(",");//未税价格
                    sw.Write(Convert.ToDateTime(drs[i]["CreateDate"].ToString()).ToString("yyMMdd")); sw.Write(",");//年简称080424 请购日期
                    sw.Write(drs[i]["CreateByName"].ToString()); sw.Write(",");//请购申请人
                    string ls = drs[i]["CreateById"].ToString() + drs[i]["CreateByName"].ToString() + "/" + drs[i]["PoVendorId"].ToString() + drs[i]["PoVendorName"].ToString();
                    sw.Write(ls.Substring(0, 20)); sw.Write(",");//请购申请人工号+姓名/采购供应商代码+名称（20个字）
                    sw.Write(drs[i]["rowid"].ToString()); sw.Write(",");  //采购行号                    
                    sw.Write("\r\n");
                }

                sw.Flush();
                sw.Close();
                fs.Close();

                sftp.Connect();
                bool bf = sftp.Put(FilePath + filename + ".txt", "/apps/OA/" + filename + ".txt");
                sftp.Disconnect();

                if (bf)
                {
                    SqlParameter[] param2 = new SqlParameter[]
                         {
                                   new SqlParameter("@flag","PO_Update"),
                                   new SqlParameter("@pono",pono)
                         };
                    SQLHelper.ExecuteNonQuery("usp_Winjob", param2);
                }

            }
            catch (Exception ex)
            {
                result = 1;
            }           

            return result;
        }

        #endregion

    }
}
