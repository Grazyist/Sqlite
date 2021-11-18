using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SQLite;
using System.Data;
using NPOI;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.UserModel;

namespace sqliteReader
{
  
    class SqliteReadToExcel
    {
        /// <summary>
        /// 从sqlite读取数据，保存到DataTabel 中，进行保存
        /// </summary>
        /// <param name="connStr"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        private static System.Data.DataTable ReadFromSqlite(string databasePath ,string table)
        {
            string connStr = @"Data Source="+databasePath;
            SQLiteConnection conn = new SQLiteConnection(connStr);
            
            try {
                conn.Open();
            }catch(SQLiteException sqlExcept)
            {
                return null;
            }
            string query = "select * from " + table;
            SQLiteCommand command = new SQLiteCommand(query,conn);

            SQLiteDataAdapter da = new SQLiteDataAdapter(command);
            System.Data.DataTable dt = new System.Data.DataTable();
            da.Fill(dt);
            
            //if(dt.Rows.Count==0|| dt.Columns.Count==0)
            //{ Console.WriteLine("没有读取到任何数据"); }
            //else
            //{ Console.WriteLine($"row:{dt.Rows.Count},col:{dt.Columns.Count}"); }

            conn.Close();
            return dt;
        }

        public static void Write2Xls(string filePath,string sheetName,string databasePath,string tableName)
        {
            DataTable dt = ReadFromSqlite(databasePath, tableName);
            IWorkbook wkBook = new HSSFWorkbook();//新建工作簿对象
            ISheet wkSheet = wkBook.CreateSheet(sheetName);//创建工作表对象

            //将datatable中的数据插入到Excel表格里面
            //向Excel中写入表头
            IRow headerRow = wkSheet.CreateRow(0);
            for(int k=0;k<dt.Columns.Count;k++)
            {
                headerRow.CreateCell(k).SetCellValue(dt.Columns[k].ColumnName);
            }

            //写入表格内容
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //创建一行
                IRow irow = wkSheet.CreateRow(i + 1);
                for(int j=0;j<dt.Columns.Count;j++)
                {
                    irow.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //把内存中的workBook对象写入到磁盘上
            

            FileStream fs = File.OpenWrite(filePath);
            try {
                wkBook.Write(fs);
                wkBook.Close();
            }catch(Exception e)
            {
                Console.WriteLine("there produce a exception"+e.Message.ToString());
            }
            fs.Close();
            fs.Dispose();
        }
    }
}
