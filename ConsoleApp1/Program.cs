# region Using Directives
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

# endregion
namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {

            /*****建立输入流的 workbookIn 对象 *****/
            IWorkbook workbookIn = null;      //新建Workbook对象
            string filename = "Y:\\赤坂亭\\工作簿1.xls";
            FileStream fileStream = new FileStream(@"Y:\\赤坂亭\\工作簿1.xls", FileMode.Open, FileAccess.Read);


            if(filename.IndexOf(".xlsx") > 0)  //判断是否为2007版本
            {
                workbookIn = new XSSFWorkbook(fileStream);    //xlsx数据写入workbookIn
            }
            else if( filename.IndexOf(".xls") > 0)      //判断是否为2003版本
            {
                workbookIn = new HSSFWorkbook(fileStream);    //xls数据写入workbookIn
            }


            /*****建立输出流的 workbookIn 对象 --- workbookOut*****/
            HSSFWorkbook workbookOut = new HSSFWorkbook();      // 建立输出流的 workbookOut 用于输出文件
            workbookOut.CreateSheet("Sheet1");                  // 建立新的表单 名称： Sheet1
            HSSFSheet sheetNew = (HSSFSheet)workbookOut.GetSheet("Sheet1");     // 获取 名称：Sheet1 的工作表


            try
            {
                ISheet sheet = workbookIn.GetSheetAt(0);  //获取第一个工作表
                IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据
                int newRow = 0;
                for (int i = 3; i <= sheet.LastRowNum; i++)  //对工作表每一行
                {
                    Console.WriteLine(i);
                    row = sheet.GetRow(i);   //row读入第i行数据
                    if (row != null)
                    {
                        for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列
                        {
                            string cellValue = row.GetCell(j).ToString(); //获取i行j列数据
                            if ((cellValue != "") && (cellValue != "TRUE"))      //获取数据 获取条件 非空 TRUE 作为截止信号
                            {
                                sheetNew.CreateRow(newRow);                  //从店铺开始 每次创建一行
                                HSSFRow sheetRow = (HSSFRow)sheetNew.GetRow(newRow);     // 获取新的行 作为对象
                                HSSFCell[] sheetCell = new HSSFCell[4];         // 每行建立四个列
                                sheetCell[0] = (HSSFCell)sheetRow.CreateCell(0);        // 建立 列[0] 用于 填写店名
                                sheetCell[1] = (HSSFCell)sheetRow.CreateCell(1);        // 建立 列[1] 用于 填写 货品名称
                                sheetCell[2] = (HSSFCell)sheetRow.CreateCell(2);        // 建立 列[2] 用于 填写 货品单位
                                sheetCell[3] = (HSSFCell)sheetRow.CreateCell(3);        // 建立 列[3] 用于 填写 货品数量

                                if (j == 0)     // 输入流 从第四行开始 每列第一位为 店铺名称   故做一个判断
                                {
                                    sheetCell[0].SetCellValue(cellValue);       //填写 店名
                                    sheetCell[1].SetCellValue("品名");          //填写 品名
                                    sheetCell[2].SetCellValue("单位");          //填写 单位
                                    sheetCell[3].SetCellValue("数量");          //填写 数量
                                    Console.WriteLine($"{cellValue} 品名 单位 数量");
                                    newRow++;       // 行数增加 进入下一行
                                }
                                else
                                { 
                                    IRow row2 = sheet.GetRow(1);        // 重新定义一个 行 迭代器 
                                    string cellName = row2.GetCell(j).ToString();       // 新的行迭代器 用于获取 商品名称     cellName 商品名称
                                    sheetCell[1].SetCellValue(cellName);        // 填写 货品名称
                                    IRow row3 = sheet.GetRow(2);        // 重新定义一个 行 迭代器
                                    string cellUnit = row3.GetCell(j).ToString();       // 新的行迭代器 用于获取 商品单位     cellUnit 商品单位
                                    sheetCell[2].SetCellValue(cellUnit);        // 填写 货品单位
                                    Double price = Convert.ToDouble(cellValue); // 字符串 转换为 32位浮点型数
                                    sheetCell[3].SetCellValue(price);       // 填写 货品数量 
                                    Console.WriteLine($"{cellName}  {cellUnit}  {cellValue}");      //控制台校验
                                    newRow++;
                                 }
                            }
                        }
                    }
                    sheetNew.CreateRow(newRow);     // 新建一行 分开每家店铺
                    newRow++;
                }
                FileStream fileOut = new FileStream(@"Y:\\赤坂亭\\test.xls", FileMode.Create);
                workbookOut.Write(fileOut);
                workbookOut.Close();
                //Console.ReadKey();
                fileStream.Close();
                workbookIn.Close();

            }
            catch(IOException ex)
            {
                Console.WriteLine(ex.StackTrace);
            }
        }

    }
}
