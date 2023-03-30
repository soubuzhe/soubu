using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace Soubu.Helper.NPOI
{
    /// <summary>
    /// 操作xls格式的Excel文件
    /// </summary>
    public class SoubuNPOI
    {
        private IWorkbook? iWorkBook; //读取时用
        private HSSFWorkbook? hssfWorkbook; //创建时用
        private string? filePath;
        #region 读取Excel文件 构造函数CommonNPOI(HSSFWorkbook workBook)
        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="workBook">工作簿对象</param>
        public SoubuNPOI(HSSFWorkbook workBook)
        {
            this.iWorkBook = workBook;
        }
        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="path">excel的绝对路径，如 C:\a.xls
        /// <para>context.Server.MapPath("../WenJianZZ/FuJian/")【转绝对路径】</para>
        /// </param>
        public SoubuNPOI(string path)
        {
            this.filePath = path;
            try
            {
                using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    if (path.ToLower().EndsWith(".xlsx"))
                        this.iWorkBook = new XSSFWorkbook(stream);
                    else
                        this.iWorkBook = new HSSFWorkbook(stream);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(path + "\r\n\r\nOffice文件读取失败，您上传的文件为非标准Office文件，请MS Office打开另存为后再上传。\r\n<br />" + ex.Message);
            }
            //finally //删除在方法里控制
            //{
            //    if (System.IO.File.Exists(this.filePath.Replace(@"\\", @"/"))) //文件存在时删除
            //    {
            //        System.IO.File.Delete(this.filePath.Replace(@"\\", @"/"));
            //    }
            //}
        }
        #endregion
        #region 无参构造方法 ，创建Excel时调用
        /// <summary>
        /// 创建Excel
        /// <para>1、CreateWorkbook</para>
        /// <para>2、CreateSheet</para>
        /// <para>3、CreateRow</para>
        /// <para>4、CreateCell【此步骤为第3步的扩展方法】</para>
        /// <para>5、SaveExcel</para>
        /// </summary>
        public SoubuNPOI() { }
        #endregion
        #region GetCellData读取Excel单元格数据
        /// <summary>
        /// 读取Excel单元格数据
        /// </summary>
        /// <param name="cell">单元格对象</param>
        /// <returns>单元格数据</returns>
        public string GetCellData(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric: //excel中日期和数字最终保存的数据类型都是int CELL_TYPE_NUMERIC
                    if (HSSFDateUtil.IsCellDateFormatted(cell)) //IsValidExcelDate：有效日期 IsInternalDateFormat：内部日期格式  IsCellInternalDateFormatted：单元格内部日期格式
                    { //日期类型时
                        try
                        { //日期格式不规范时读取失败
                            return cell.DateCellValue.ToString("yyyy-MM-dd");
                        }
                        catch
                        {
                            return cell.NumericCellValue.ToString();
                        }
                    }
                    else //cell.CellType==CellType.NUMERIC
                    { //数字类型时
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.String: //字符串类型CELL_TYPE_STRING
                    return cell.StringCellValue;
                case CellType.Formula: //公式类型 CELL_TYPE_FORMULA
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.Boolean:
                            return cell.BooleanCellValue.ToString();
                        case CellType.Numeric:
                            return cell.NumericCellValue.ToString();
                        case CellType.String:
                            return cell.StringCellValue;
                        default:
                            return "公式类型数据获取错误";
                    }
                //break;
                case CellType.Blank: //空数据类型 CELL_TYPE_BLANK
                    return string.Empty;
                default:
                    return string.Empty;
            }
        }
        #endregion
        #region ReadExcel读取excel内容
        /// <summary>
        /// 读取excel内容【去除字符两边的空格】
        /// <para>key：工作簿序号，value：列值集合【每个List就是一行】</para>
        /// <para>返回值只有单元格内容</para>
        /// </summary>
        /// <param name="bDelete">成功或错误时是否删除文件，true：删除，false：不删除</param>
        /// <param name="listSheetName">工作表的名称，下标与返回值Dictionary的key对应，Dictionary的key就是工作表的序号</param>
        /// <returns>key：工作表序号，value：行列值集合
        /// <para>每个List&lt;string&gt;就是一行，List&lt;List&lt;SoubuNPOICell_M&gt;&gt;为列对象</para>
        /// </returns>
        public Dictionary<int, List<List<string>>> ReadExcel(bool bDelete, out List<string> listSheetName)
        {
            Dictionary<int, List<List<string>>> dicCellStr;
            Dictionary<int, List<List<SoubuNPOICell_M>>> dicCellM;
            ReadExcel(bDelete, false, out listSheetName, out dicCellStr, out dicCellM, true);
            return dicCellStr;
        }
        /// <summary>
        /// 读取excel内容
        /// <para>key：工作簿序号，value：列值集合【每个List就是一行】</para>
        /// <para>返回值只有单元格内容</para>
        /// </summary>
        /// <param name="bDelete">成功或错误时是否删除文件，true：删除，false：不删除</param>
        /// <param name="listSheetName">工作表的名称，下标与返回值Dictionary的key对应，Dictionary的key就是工作表的序号</param>
        /// <param name="bTrim">是否去除两边的空格，true：去除，false：不去除</param>
        /// <returns>key：工作表序号，value：行列值集合
        /// <para>每个List&lt;string&gt;就是一行，List&lt;List&lt;SoubuNPOICell_M&gt;&gt;为列对象</para>
        /// </returns>
        public Dictionary<int, List<List<string>>> ReadExcel(bool bDelete, out List<string> listSheetName, bool bTrim)
        {
            Dictionary<int, List<List<string>>> dicCellStr;
            Dictionary<int, List<List<SoubuNPOICell_M>>> dicCellM;
            ReadExcel(bDelete, false, out listSheetName, out dicCellStr, out dicCellM, bTrim);
            return dicCellStr;
        }
        /// <summary>
        /// 读取excel内容【去除字符两边的空格】
        /// <para>key：工作簿序号，value：列值集合【每个List就是一行】</para>
        /// <para>返回值包含单元格属性</para>
        /// </summary>
        /// <param name="listSheetName">工作表的名称，下标与返回值Dictionary的key对应，Dictionary的key就是工作表的序号</param>
        /// <param name="bDelete">成功或错误时是否删除文件，true：删除，false：不删除</param>
        /// <returns>key：工作表序号，value：行列值集合
        /// <para>每个List&lt;string&gt;就是一行，List&lt;List&lt;SoubuNPOICell_M&gt;&gt;为列对象</para>
        /// </returns>
        public Dictionary<int, List<List<SoubuNPOICell_M>>> ReadExcel(out List<string> listSheetName, bool bDelete)
        {
            Dictionary<int, List<List<string>>> dicCellStr;
            Dictionary<int, List<List<SoubuNPOICell_M>>> dicCellM;
            ReadExcel(bDelete, true, out listSheetName, out dicCellStr, out dicCellM, true);
            return dicCellM;
        }
        /// <summary>
        /// 读取excel内容
        /// <para>key：工作簿序号，value：列值集合【每个List就是一行】</para>
        /// <para>返回值包含单元格属性</para>
        /// </summary>
        /// <param name="listSheetName">工作表的名称，下标与返回值Dictionary的key对应，Dictionary的key就是工作表的序号</param>
        /// <param name="bDelete">成功或错误时是否删除文件，true：删除，false：不删除</param>
        /// <param name="bTrim">是否去除两边的空格，true：去除，false：不去除</param>
        /// <returns>key：工作表序号，value：行列值集合
        /// <para>每个List&lt;string&gt;就是一行，List&lt;List&lt;SoubuNPOICell_M&gt;&gt;为列对象</para>
        /// </returns>
        public Dictionary<int, List<List<SoubuNPOICell_M>>> ReadExcel(out List<string> listSheetName, bool bDelete, bool bTrim)
        {
            Dictionary<int, List<List<string>>> dicCellStr;
            Dictionary<int, List<List<SoubuNPOICell_M>>> dicCellM;
            ReadExcel(bDelete, true, out listSheetName, out dicCellStr, out dicCellM, bTrim);
            return dicCellM;
        }
        /// <summary>
        /// 读取excel内容
        /// <para>key：工作簿序号，value：列值集合【每个List就是一行】</para>
        /// </summary>
        /// <param name="bDelete">成功或错误时是否删除文件，true：删除，false：不删除</param>
        /// <param name="bFlag">true:单元格内容是合并信息(以实体方式返回)，false:只返回单元格内的内容(以string方式返回)</param>
        /// <param name="listSheetName">工作表的名称，下标与返回值Dictionary的key对应，Dictionary的key就是工作表的序号</param>
        /// <param name="dicCellStr">excel集合</param>
        /// <param name="dicCellM">excel集合</param>
        /// <param name="bTrim">是否去除两边的空格，true：去除，false：不去除</param>
        /// <returns>key：工作表序号，value：行列值集合
        /// <para>每个List&lt;string&gt;就是一行，List&lt;List&lt;SoubuNPOICell_M&gt;&gt;为列对象</para>
        /// </returns>
        private void ReadExcel(bool bDelete, bool bFlag, out List<string> listSheetName, out Dictionary<int, List<List<string>>> dicCellStr, out Dictionary<int, List<List<SoubuNPOICell_M>>> dicCellM, bool bTrim)
        {
            int bookNum = 0, rowNum = 0, cellNum = 0; //行号、列号
            listSheetName = new List<string>();
            try
            {
                dicCellM = new Dictionary<int, List<List<SoubuNPOICell_M>>>();
                dicCellStr = new Dictionary<int, List<List<string>>>();
                ISheet sheet;
                IRow row;
                ICell cell;
                if (this.iWorkBook == null) throw new Exception(nameof(SoubuNPOI)+ "的ReadExcel方法出现错误，原因：iWorkBook==null");
                for (int i = 0; i < this.iWorkBook.NumberOfSheets; i++) //循环工作簿
                {
                    bookNum = i;
                    listSheetName.Add(iWorkBook.GetSheetName(i));
                    sheet = this.iWorkBook.GetSheetAt(i);
                    row = sheet.GetRow(0);
                    if (row == null) continue;
                    List<List<SoubuNPOICell_M>> listRow_M = new List<List<SoubuNPOICell_M>>();
                    List<List<string>> listRow_Str = new List<List<string>>();
                    for (int x = 0; x <= sheet.LastRowNum; x++)//从第一行开始读取【循环行】
                    {
                        rowNum = x;
                        row = sheet.GetRow(x);
                        if (row == null) continue;
                        List<SoubuNPOICell_M> listCell_M = new List<SoubuNPOICell_M>();
                        List<string> listCell_Str = new List<string>();
                        for (int y = 0; y < row.LastCellNum; y++) //循环列
                        {
                            cellNum = y;
                            cell = row.GetCell(y);
                            Point start, end;
                            string Cellval = cell == null ? string.Empty : (bTrim ? GetCellData(cell).Trim() : GetCellData(cell));
                            if (bFlag)
                            {
                                bool isMergeCell = IsMergeCell(sheet, x, y, out start, out end);
                                SoubuNPOICell_M m = new SoubuNPOICell_M();
                                m.CellNum = y;
                                m.Content = Cellval;
                                m.EndMerge = end;
                                m.IsMerge = isMergeCell;
                                m.RowNum = x;
                                m.StartMerge = start;
                                listCell_M.Add(m);
                            }
                            else
                            {
                                listCell_Str.Add(Cellval);
                            }
                        }
                        if (bFlag)
                        {
                            listRow_M.Add(listCell_M);
                        }
                        else
                        {
                            listRow_Str.Add(listCell_Str);
                        }
                    }
                    if (bFlag)
                    {
                        dicCellM.Add(i, listRow_M);
                    }
                    else
                    {
                        dicCellStr.Add(i, listRow_Str);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("CommonNPOI.ReadExcel方法异常，《{0}》第{1}行，第{2}列，错误原因：{3}", listSheetName[bookNum], rowNum + 1, cellNum + 1, ex.ToString()));
            }
            finally
            {
                if (bDelete && System.IO.File.Exists(this.filePath?.Replace(@"\\", @"/"))) //文件存在时删除
                {
                    System.IO.File.Delete(this.filePath.Replace(@"\\", @"/"));
                }
            }
        }
        #endregion
        #region IsMergeCell获取当前单元格所在的合并单元格的位置
        /// <summary>
        /// 获取当前单元格所在的合并单元格的位置
        /// </summary>
        /// <param name="sheet">sheet表单</param>
        /// <param name="rowIndex">行索引 0开始</param>
        /// <param name="colIndex">列索引 0开始</param>
        /// <param name="start">合并单元格左上角坐标，2,2表示C3单元格</param>
        /// <param name="end">合并单元格右下角坐标，7,3表示D8单元格</param>
        /// <returns>返回false表示非合并单元格</returns>
        private bool IsMergeCell(ISheet sheet, int rowIndex, int colIndex, out Point start, out Point end)
        {
            bool result = false;
            start = new Point(0, 0);
            end = new Point(0, 0);
            if ((rowIndex < 0) || (colIndex < 0)) return result;
            int regionsCount = sheet.NumMergedRegions;
            for (int i = 0; i < regionsCount; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                //sheet.IsMergedRegion(range); 
                if (rowIndex >= range.FirstRow && rowIndex <= range.LastRow && colIndex >= range.FirstColumn && colIndex <= range.LastColumn)
                {
                    start = new Point(range.FirstRow, range.FirstColumn);
                    end = new Point(range.LastRow, range.LastColumn);
                    result = true;
                    break;
                }
            }
            return result;
        }
        #endregion
        //#region CreateColumn创建列（创建单元格）
        ///// <summary>
        ///// 创建列（创建单元格）字符串[string]
        ///// </summary>
        ///// <param name="row">创建单元格行对象（在第几行创建单元格）</param>
        ///// <param name="cellNum">创建第几列单元格</param>
        ///// <param name="cellValue">赋给单元格的值</param>
        ///// <param name="cellStyle">单元格样式</param>
        //public void CreateColumn(HSSFRow row, int cellNum, string cellValue, ICellStyle cellStyle)
        //{
        //    ICell cell = row.CreateCell(cellNum, CellType.String);
        //    //单元格的数据类型
        //    //0：数字或日期（HSSFCell.CELL_TYPE_NUMERIC）
        //    //1：字符串（HSSFCell.CELL_TYPE_STRING）
        //    //2：公式（HSSFCell.CELL_TYPE_FORMULA）
        //    //3：空白（HSSFCell.CELL_TYPE_BLANK）
        //    //4：布尔（HSSFCell.CELL_TYPE_BOOLEAN）
        //    //5：错误（HSSFCell.CELL_TYPE_ERROR）
        //    cell.SetCellValue(cellValue);
        //    cell.CellStyle = cellStyle;
        //}
        ///// <summary> 
        ///// 创建列（创建单元格） 数字[int]
        ///// </summary>
        ///// <param name="row">创建单元格行对象（在第几行创建单元格）</param>
        ///// <param name="cellNum">创建第几列单元格</param>
        ///// <param name="cellValue">赋给单元格的值</param>
        ///// <param name="cellStyle">单元格样式</param>
        //public void CreateColumn(HSSFRow row, int cellNum, int cellValue, ICellStyle cellStyle)
        //{
        //    ICell cell = row.CreateCell(cellNum, CellType.Numeric);
        //    cell.SetCellValue(cellValue);
        //    cell.CellStyle = cellStyle;
        //}
        //#endregion

        #region 1、CreateWorkbook创建Excel文件并返回Excel对象
        /// <summary>
        /// 1、创建Excel文件并返回Excel对象
        /// </summary>
        public HSSFWorkbook CreateWorkbook()
        {
            //Create 一个Excel对象
            this.hssfWorkbook = new HSSFWorkbook();
            //Create Excel的属性中的来源以及说明等
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = string.Empty;// 创建Excel文件的公司名称
            dsi.Category = string.Empty; //文件的类别
            dsi.Manager = string.Empty; //创建Excel文件的姓名
            //创建好的对象赋给hssfWorkbook,这样才能保证这些信息被写入文件
            this.hssfWorkbook.DocumentSummaryInformation = dsi;
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = string.Empty; //Excel文件主题
            si.Title = string.Empty; //Excel文件标题
            si.ApplicationName = string.Empty; //创建Excel文件的应用程序名称
            si.Author = string.Empty; //作者
            si.LastAuthor = string.Empty; //最后一次修改的作者
            si.Comments = "NPOI创建"; //备注
            si.CreateDateTime = DateTime.Now; //创建时间
            //创建好的对象赋给hssfWorkbook,这样才能保证这些信息被写入文件
            this.hssfWorkbook.SummaryInformation = si;
            return this.hssfWorkbook;
        }
        #endregion
        #region 2、CreateSheet创建工作表并返回工作表对象
        /// <summary>
        /// 2、创建工作表并返回工作表对象
        /// </summary>
        /// <param name="sheet">工作表的名称，如：Sheet1
        /// <para>传空时，默认为Sheet1</para>
        /// </param>
        public ISheet? CreateSheet(string sheet)
        {
            return hssfWorkbook?.CreateSheet(string.IsNullOrWhiteSpace(sheet) ? "Sheet1" : sheet);
        }
        #endregion
        #region 3、CreateColumn创建行并返回行对象
        /// <summary>
        /// 3、创建行并返回行对象
        /// </summary>
        /// <param name="sheet">工作表对象，如：Sheet1</param>
        /// <param name="rowNum">创建的行号</param>
        /// <param name="cellNum">列号，当前行需要创建的列号
        /// <para>cellNum.Count、cellVal.Count、cellSytle.Count必须相等，否则只返回行对象，忽略创建列和忽略写入列值</para>
        /// </param>
        /// <param name="cellVal">列值，当前行需要填入的列值</param>
        /// <param name="cellSytle">一组单元格的样式，List中的元素值可以为NULL，某个列的样式可以不设置
        /// <para>创建样式 ICellStyle cellStyle = hssfWorkbook.CreateCellStyle();  </para>
        /// <para> 设置单元格上边框线 cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;</para>
        /// <para>文字水平对齐方式 cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; </para>
        /// <para>文字垂直对齐方式 cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; </para>
        /// <para>是否换行 cellStyle.WrapText = true;  </para>
        /// <para>缩小字体填充 cellStyle.ShrinkToFit = true; </para>
        /// </param>
        /// <returns>返回行对象</returns>
        public IRow CreateRow(ISheet sheet, int rowNum, List<int> cellNum, List<string> cellVal, List<ICellStyle> cellSytle)
        {
            IRow row = sheet.CreateRow(rowNum);
            if (cellNum == null || cellVal == null || cellSytle == null || cellNum.Count == 0 || cellVal.Count == 0 || cellSytle.Count == 0) return row;
            if (cellNum.Count != cellVal.Count || cellNum.Count != cellSytle.Count) return row;
            for (int i = 0; i < cellNum.Count; i++)
            {
                ICell cell = row.CreateCell(cellNum[i], CellType.String);
                cell.SetCellValue(cellVal[i]);
                if (cellSytle != null) cell.CellStyle = cellSytle[i];
            }
            return row;
        }
        /// <summary>
        /// 3、创建行并返回行对象
        /// </summary>
        /// <param name="sheet">工作表对象，如：Sheet1</param>
        /// <param name="rowNum">创建的行号</param>
        /// <param name="cellNum">列号，当前行需要创建的列号
        /// <para>cellNum.Count、cellVal.Count、cellSytle.Count必须相等，否则只返回行对象，忽略创建列和忽略写入列值</para>
        /// </param>
        /// <param name="cellVal">列值，当前行需要填入的列值</param>
        /// <param name="cellSytle">一组单元格的样式，List中的元素值可以为NULL，某个列的样式可以不设置
        /// <para>创建样式 ICellStyle cellStyle = hssfWorkbook.CreateCellStyle();  </para>
        /// <para> 设置单元格上边框线 cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;</para>
        /// <para>文字水平对齐方式 cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; </para>
        /// <para>文字垂直对齐方式 cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; </para>
        /// <para>是否换行 cellStyle.WrapText = true;  </para>
        /// <para>缩小字体填充 cellStyle.ShrinkToFit = true; </para>
        /// </param>
        /// <returns>返回行对象</returns>
        public IRow CreateRow(ISheet sheet, int rowNum, List<int> cellNum, List<int> cellVal, List<ICellStyle> cellSytle)
        {
            IRow row = sheet.CreateRow(rowNum);
            if (cellNum == null || cellVal == null || cellSytle == null || cellNum.Count == 0 || cellVal.Count == 0 || cellSytle.Count == 0) return row;
            if (cellNum.Count != cellVal.Count || cellNum.Count != cellSytle.Count) return row;
            for (int i = 0; i < cellNum.Count; i++)
            {
                ICell cell = row.CreateCell(cellNum[i], CellType.Numeric);
                cell.SetCellValue(cellVal[i]);
                if (cellSytle != null) cell.CellStyle = cellSytle[i];
            }
            return row;
        }
        #endregion
        #region 4、CreateCell创建列，可以忽略，为第3步的扩展方法
        /// <summary>
        /// 4、创建列，可以忽略，为第3步的扩展方法
        /// </summary>
        /// <param name="row">row对象【由CreateRow方法返回】</param>
        /// <param name="cellNum">单元格列号</param>
        /// <param name="cellVal">单元格列值</param>
        /// <param name="cellSytle">一组单元格的样式，List中的元素值可以为NULL，某个列的样式可以不设置
        /// <para>创建样式 ICellStyle cellStyle = hssfWorkbook.CreateCellStyle();  </para>
        /// <para> 设置单元格上边框线 cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;</para>
        /// <para>文字水平对齐方式 cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; </para>
        /// <para>文字垂直对齐方式 cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; </para>
        /// <para>是否换行 cellStyle.WrapText = true;  </para>
        /// <para>缩小字体填充 cellStyle.ShrinkToFit = true; </para>
        /// </param>
        public ICell CreateCell(IRow row, int cellNum, string cellVal, ICellStyle cellSytle)
        {
            ICell cell = row.CreateCell(cellNum, CellType.String);
            cell.SetCellValue(cellVal);
            if (cellSytle != null) cell.CellStyle = cellSytle;
            return cell;
        }
        /// <summary>
        /// 4、创建列，可以忽略，为第3步的扩展方法
        /// </summary>
        /// <param name="row">row对象【由CreateRow方法返回】</param>
        /// <param name="cellNum">单元格列号</param>
        /// <param name="cellVal">单元格列值</param>
        /// <param name="cellSytle">一组单元格的样式，List中的元素值可以为NULL，某个列的样式可以不设置
        /// <para>创建样式 ICellStyle cellStyle = hssfWorkbook.CreateCellStyle();  </para>
        /// <para> 设置单元格上边框线 cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;</para>
        /// <para>文字水平对齐方式 cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; </para>
        /// <para>文字垂直对齐方式 cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; </para>
        /// <para>是否换行 cellStyle.WrapText = true;  </para>
        /// <para>缩小字体填充 cellStyle.ShrinkToFit = true; </para>
        /// </param>
        public ICell CreateCell(IRow row, int cellNum, int cellVal, ICellStyle cellSytle)
        {
            ICell cell = row.CreateCell(cellNum, CellType.Numeric);
            cell.SetCellValue(cellVal);
            if (cellSytle != null) cell.CellStyle = cellSytle;
            return cell;
        }
        #endregion
        #region 5、SaveExcel()保存Excel文件
        /// <summary>
        /// 5、保存Excel文件
        /// </summary>
        /// <param name="savePath">绝对路径，如：C:\工资条.xls</param>
        public void SaveExcel(string savePath)
        {
            //创建文件并写入
            try
            {
                using (FileStream file = new FileStream(savePath, FileMode.Create))
                {
                    this.hssfWorkbook?.Write(file);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Excel文件创建失败，导出失败\r\n<br />" + ex.Message);
            }
        }
        #endregion
        #region SetColumnWidth设置单元格的宽度
        /// <summary>
        /// 设置单元格的宽度
        /// </summary>
        /// <param name="sheet">工作表对象</param>
        /// <param name="column">列对象</param>
        /// <param name="width">宽，单位：1个半角字符（设为2时，实际宽1.29）</param>
        public void SetColumnWidth(ISheet sheet, int column, double width)
        {
            sheet.SetColumnWidth(column, (int)((width + 0.72) * 256));
        }
        #endregion
        #region JoinRowCol合并Excel的行或列
        /// <summary>
        /// 合并Excel的行或列
        /// </summary>
        /// <param name="sheet">待合并行列的excel工作表对象</param>
        /// <param name="firstRow">从当前行开始合并（0代表第1行）</param>
        /// <param name="lastRow">到当前行结束合并（0代表第1行）</param>
        /// <param name="firstCol">从当前列开始合并（0代表A列）</param>
        /// <param name="lastCol">到当前列结束合并（0代表A列）</param>
        public void JoinRowCol(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            sheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol)); //合并第N行的Q到T列
        }
        #endregion
    }
}
