#region
/* ==============================================================================
  * 功能描述：PKExcel  
  * 创 建 者：GuoLiang
  * 创建日期：2012-12-12 14:45:05
  * CLR 版本：4.0.30319.18010
  * ==============================================================================*/
#endregion

using System;
using System.Data.OleDb;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace PLANCK.ExcelHelper
{
    public class PKExcel
    {
        public static void GridToExcel(DataTable table, string[] header)
        {
            try
            {
                Application excel = new Application ();
                excel.Application.Workbooks.Add (true);
                Parallel.For (0, header.Length, i => excel.Cells [1, i + 1] = header [i]);
                Parallel.For (0, table.Rows.Count, i => Parallel.For (0, table.Columns.Count,
                                                                      j =>
                                                                          {
                                                                              excel.Cells [i + 2, j + 1] = "'" +
                                                                                                           table.Rows [i
                                                                                                               ] [j];
                                                                          }
                                                            ));
                excel.Columns.AutoFit ();
                excel.Columns.WrapText = true;
                excel.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "信息失败！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static DataTable GetExcleToTable()
        {
            try
            {
                OpenFileDialog openFileDialog1 =new OpenFileDialog ();
                openFileDialog1.Filter = "*.xlsx|*.xlsx";
                openFileDialog1.FileName = "*.xlsx";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string path = openFileDialog1.FileName;
                    string strConn ="Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + path +";Extended Properties='Excel 12.0; HDR=NO; IMEX=1'";
                    using (OleDbConnection conn = new OleDbConnection(strConn))
                    {
                        const string sql = "SELECT * FROM [Sheet1$]";
                        OleDbDataAdapter adp = new OleDbDataAdapter(sql, conn);
                        DataSet ds = new DataSet();
                        adp.Fill(ds, "0");
                        return ds.Tables[0];
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "信息失败！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return null;
        }
    } //End  PKExcel
} //End  PLANCK.ExcelHelper
