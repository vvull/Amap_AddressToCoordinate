using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            var dt = ExcelToDataTable(@"C:\Users\ZengJW\Desktop\站点地址.xls", true);
            List<string> address = new List<string>();//地址集合
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                address.Add(dt.Rows[i][1].ToString());
            }
            if (DataTableToExcel(AddressToAmap(address.ToArray()), @"C:\Users\ZengJW\Desktop\test.xls"))
            {
                Console.WriteLine("ok！");
            }


        }

        /// <summary>
        /// 将地址转化为高德地图坐标（云图数据模板）
        /// </summary>
        /// <param name="str">地址</param>
        /// <returns>返回表格，列1为name，列2为address，列3为经度，列4为纬度，列5为telephone</returns>
        public static DataTable AddressToAmap(params string[] address)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("name");
            dt.Columns.Add("address");
            dt.Columns.Add("x");
            dt.Columns.Add("y");
            dt.Columns.Add("telephone");
            StringBuilder sb = new StringBuilder("https://restapi.amap.com/v3/place/text?s=rsv3&key=22c0612566037e2324502ae619a8f4a9&page=1&offset=10&city=440600&language=zh_cn&callback=jsonp_585151_&platform=JS&logversion=2.0&sdkversion=1.3&appname=https://lbs.amap.com/console/show/picker&csid=1FCC8F20-A9E7-4E79-81F4-CE27742A797C&keywords=");
            HttpClient hc = new HttpClient();
            for (int i = 0; i < address.Length; i++)
            {

                //把需要查询的地址拼接到url
                sb.Append(address[i]);
                string queryStr = sb.ToString();
                var responseMessage = hc.GetAsync(queryStr);
                var content = responseMessage.Result.Content.ReadAsStringAsync().Result;

                //去掉地址(查询字符串中"keywords="后面的地址)，重复使用
                sb.Remove(queryStr.LastIndexOf('=') + 1, queryStr.Length - (queryStr.LastIndexOf('=') + 1));

                //去掉多余字符串，保留核心数据
                dynamic data = JsonConvert.DeserializeObject(content.Substring(content.IndexOf('(') + 1, content.Length - content.IndexOf('(') - 1 - 1));

                //取搜索结果第一项
                if (data.pois == null || data.pois.Count == 0)
                {
                    continue;//给定地址搜索结果为空，跳过
                }
                dynamic poi = data.pois[0];
                string[] location = poi.location.ToString().Split(',');//分割经纬度
                string tel = "00000000000";
                if (string.IsNullOrEmpty(poi.tel.ToString()))
                {
                    tel.Replace(':', ',');//tel不为空时
                }

                //存储poi数据
                dt.Rows.Add(poi.name, poi.address, location[0], location[1], tel);
            }
            hc.Dispose();
            return dt;
        }

        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="isColumnName">第一行是否是列名</param>
        /// <returns>返回datatable</returns>
        public static DataTable ExcelToDataTable(string filePath, bool isColumnName)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);
                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行
                                int cellCount = firstRow.LastCellNum;//列数
                                                                     //构建datatable的列
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }
                                //填充行
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null) continue;
                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        /// <summary>
        /// 将datatable导入到excel
        /// </summary>
        /// <param name="dt">需要导入的数据</param>
        /// <returns>导入结果</returns>
        public static bool DataTableToExcel(DataTable dt, string filePath)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new HSSFWorkbook();
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数
                                                       //设置列头
                    row = sheet.CreateRow(0);//excel第一行设为列头
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }
                    //设置每行每列的单元格,
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    using (fs = File.OpenWrite(filePath))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据
                        result = true;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return false;
            }
        }

    }
}
