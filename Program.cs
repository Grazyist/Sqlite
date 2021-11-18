using System;


namespace sqliteReader
{
    class Program
    {
        
        public static void Main(string[] args)
        {
            string filePath = @"../t.xls";
            string DatabasePath = @"C:\WPF\Console\sqliteReader\db.sqlite3";
            string sheetName = "Devil";
            string tableName = "info";
            SqliteReadToExcel.Write2Xls(filePath,sheetName, DatabasePath,tableName);
        }

        //static void Write2Excel()
        //{
        //   List<Person> list = new List<Person>() {
        //        new Person(){Name="tom",Age="35",Email="ddd@qq.com" },
        //        new Person(){Name="week",Age="16",Email="456@qq.com" },
        //        new Person(){Name="tick",Age="17",Email="789@qq.com" }
        //        };
        //    //1、创建工作簿对象
        //    IWorkbook wkBook = new HSSFWorkbook();
        //    //2、在该工作簿中创建工作表对象
        //    ISheet sheet = wkBook.CreateSheet("人员信息"); 
        //    for (int i = 0; i < list.Count; i++)
        //    {
        //        //在Sheet中插入创建一行
        //        IRow row = sheet.CreateRow(i);
               
        //        row.CreateCell(0).SetCellValue(list[i].Name); //给单元格设置值：第一个参数(第几个单元格)；第二个参数(给当前单元格赋值)
        //        row.CreateCell(1).SetCellValue(list[i].Age);
        //        row.CreateCell(2).SetCellValue(list[i].Email);
        //    }
        //    //3、写入，把内存中的workBook对象写入到磁盘上
            
        //    FileStream fsWrite = File.OpenWrite(@"C:\WPF\a.xls");  //导出时Excel的文件名
        //    wkBook.Write(fsWrite);
        //    Console.WriteLine("写入成功！", "提示");
        //    fsWrite.Close(); //关闭文件流
        //    wkBook.Close();  //关闭工作簿
        //    fsWrite.Dispose(); //释放文件流
        //}
    }

}
