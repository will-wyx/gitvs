using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExpressMaster
{
    public class ExcelHelper
    {
        const int
            A = 0, B = 1, C = 2, D = 3, E = 4, F = 5, G = 6,
            H = 7, I = 8, J = 9, K = 10, L = 11, M = 12, N = 13,
            O = 14, P = 15, Q = 16, R = 17, S = 18, T = 19,
            U = 20, V = 21, W = 22, X = 23, Y = 24, Z = 25;

        public List<Data4Cfg> Data { get; set; }
        public List<ExpressEntity> DataOrder = new List<ExpressEntity>();
        public int DataOrderUnBindCount = 0;

        /// <summary>
        /// 纸单导出文件名
        /// </summary>
        public string filename = "";

        /// <summary>
        /// 加载单号
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public List<ExpressEntity> LoadOrderNumber(FileInfo file)
        {
            filename = file.Name;
            List<ExpressEntity> result = new List<ExpressEntity>();
            /* 单类型 */
            FileStream fromFile = file.OpenRead();
            XSSFWorkbook workbook = new XSSFWorkbook(fromFile);
            ISheet sheetAt = workbook.GetSheetAt(0);
            for (int i = sheetAt.FirstRowNum; i <= sheetAt.LastRowNum; ++i)
            {
                IRow row = sheetAt.GetRow(i);
                ICell onc = row.GetCell(0);
                string orderNumber = "";
                if (onc.CellType == CellType.Numeric)
                    orderNumber = onc.NumericCellValue.ToString();
                else
                    orderNumber = onc.StringCellValue;
                result.Add(new ExpressEntity
                {
                    OrderNumber = orderNumber
                });
            }
            fromFile.Close();
            DataOrder.AddRange(result);
            DataOrderUnBindCount++;
            return result;
        }

        /// <summary>
        /// 纸质单
        /// </summary>
        /// <param name="fromFile"></param>
        public void ProcessExcelTemplateC(FileStream fromFile)
        {

            XSSFWorkbook workbook = new XSSFWorkbook(fromFile);
            ISheet sheetAt = workbook.GetSheetAt(0);
            for (int i = sheetAt.FirstRowNum; i < sheetAt.LastRowNum; ++i)
            {
                IRow row = sheetAt.GetRow(i);
                /*
                0: 空
                1, 2: 表头
                */
                if (i >= 3)
                {
                    ICell
                        cellOrderNumber = row.GetCell(6) /* G列 单号 */;
                    string orderNumber = cellOrderNumber.StringCellValue;

                    ExpressEntity ee = DataOrder.Find(p => p.OrderNumber.Equals(orderNumber));
                    if (ee != null)
                    {
                        ee.OrderNumber = orderNumber;

                        ICell
                            cellWeight = row.GetCell(5) /* F列 重量 */,
                            cellDate = row.GetCell(2) /* C列 揽件日期 */,
                            cellCity = row.GetCell(8) /* I列 城市 */;

                        double weight = cellWeight.NumericCellValue;

                        double date = cellDate.NumericCellValue;
                        string city = cellCity.StringCellValue;

                        ee.Weight = weight;
                        ee.Date = date;
                        ee.City = city;
                        ee.Flag = true;
                        DataOrderUnBindCount--;
                    }
                }
            }
        }

        /// <summary>
        /// 纸单
        /// </summary>
        /// <param name="toFile"></param>
        internal void SaveExcelC(FileStream toFile)
        {
            IWorkbook toWorkbook = new XSSFWorkbook();
            ISheet toSheet = toWorkbook.CreateSheet("Sheet1");

            ICellStyle toStyleHead = toWorkbook.CreateCellStyle();
            toStyleHead.BorderBottom = BorderStyle.Thin;
            toStyleHead.BorderLeft = BorderStyle.Thin;
            toStyleHead.BorderRight = BorderStyle.Thin;
            toStyleHead.BorderTop = BorderStyle.Thin;
            toStyleHead.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
            toStyleHead.FillPattern = FillPattern.SolidForeground;

            ICellStyle toStyleGeneric = toWorkbook.CreateCellStyle();
            toStyleGeneric.BorderBottom = BorderStyle.Thin;
            toStyleGeneric.BorderLeft = BorderStyle.Thin;
            toStyleGeneric.BorderRight = BorderStyle.Thin;
            toStyleGeneric.BorderTop = BorderStyle.Thin;

            ICellStyle toStyleNull = toWorkbook.CreateCellStyle();
            toStyleNull.BorderBottom = BorderStyle.Thin;
            toStyleNull.BorderLeft = BorderStyle.Thin;
            toStyleNull.BorderRight = BorderStyle.Thin;
            toStyleNull.BorderTop = BorderStyle.Thin;
            toStyleNull.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            toStyleNull.FillPattern = FillPattern.SolidForeground;

            ICellStyle toStyleZero = toWorkbook.CreateCellStyle();
            toStyleZero.BorderBottom = BorderStyle.Thin;
            toStyleZero.BorderLeft = BorderStyle.Thin;
            toStyleZero.BorderRight = BorderStyle.Thin;
            toStyleZero.BorderTop = BorderStyle.Thin;
            toStyleZero.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
            toStyleZero.FillPattern = FillPattern.SolidForeground;

            ICellStyle toStyleDate = toWorkbook.CreateCellStyle();
            toStyleDate.BorderBottom = BorderStyle.Thin;
            toStyleDate.BorderLeft = BorderStyle.Thin;
            toStyleDate.BorderRight = BorderStyle.Thin;
            toStyleDate.BorderTop = BorderStyle.Thin;
            IDataFormat toFormatDate = toWorkbook.CreateDataFormat();
            toStyleDate.DataFormat = toFormatDate.GetFormat("yyyy-m-d");

            ICellStyle toStyleTotalAmount = toWorkbook.CreateCellStyle();
            toStyleTotalAmount.BorderBottom = BorderStyle.Thin;
            toStyleTotalAmount.BorderLeft = BorderStyle.Thin;
            toStyleTotalAmount.BorderRight = BorderStyle.Thin;
            toStyleTotalAmount.BorderTop = BorderStyle.Thin;

            ICellStyle colorStyle = toWorkbook.CreateCellStyle();
            colorStyle.BorderBottom = BorderStyle.Thin;
            colorStyle.BorderLeft = BorderStyle.Thin;
            colorStyle.BorderRight = BorderStyle.Thin;
            colorStyle.BorderTop = BorderStyle.Thin;
            colorStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
            colorStyle.FillPattern = FillPattern.SolidForeground;

            IDataFormat toFormatTotalAmount = toWorkbook.CreateDataFormat();
            toStyleTotalAmount.DataFormat = toFormatTotalAmount.GetFormat("#,##0.00");

            int toSheetRowIndex = 0;


            IRow toRow = toSheet.CreateRow(toSheetRowIndex++);
            ICell
                cellHeadIndex = toRow.CreateCell(A),    // 序号
                cellHeadOrderNumber = toRow.CreateCell(B),  // 单号
                cellHeadCity = toRow.CreateCell(C), // 城市
                cellHeadWeight = toRow.CreateCell(D),   // 重量
                cellHeadTotalAmount = toRow.CreateCell(E),  // 总金额
                cellHeadDate = toRow.CreateCell(F), // 日期
                cellHeadFirstWeight = toRow.CreateCell(G),  // 首重重量
                cellHeadFirstAmount = toRow.CreateCell(H),  // 小件首重金额
                cellHeadFirstAmountB = toRow.CreateCell(I), // 大件首重金额
                cellHeadOtherAmount = toRow.CreateCell(J);  // 续重金额

            cellHeadIndex.SetCellValue("序号");
            cellHeadOrderNumber.SetCellValue("运单号");
            cellHeadCity.SetCellValue("城市");
            cellHeadWeight.SetCellValue("重量");
            cellHeadTotalAmount.SetCellValue("金额");
            cellHeadDate.SetCellValue("日期");
            cellHeadFirstWeight.SetCellValue("首重重量");
            cellHeadFirstAmount.SetCellValue("小件首重金额");
            cellHeadFirstAmountB.SetCellValue("大件首重金额");
            cellHeadOtherAmount.SetCellValue("续重金额");

            cellHeadIndex.CellStyle = toStyleHead;
            cellHeadOrderNumber.CellStyle = toStyleHead;
            cellHeadCity.CellStyle = toStyleHead;
            cellHeadWeight.CellStyle = toStyleHead;
            cellHeadTotalAmount.CellStyle = toStyleHead;
            cellHeadDate.CellStyle = toStyleHead;
            cellHeadFirstWeight.CellStyle = toStyleHead;
            cellHeadFirstAmount.CellStyle = toStyleHead;
            cellHeadFirstAmountB.CellStyle = toStyleHead;
            cellHeadOtherAmount.CellStyle = toStyleHead;

            int index = 1;
            foreach (ExpressEntity ee in DataOrder)
            {
                IRow toItemRow = toSheet.CreateRow(toSheetRowIndex++);
                ICell
                    cellToIndex = toItemRow.CreateCell(A),
                    cellToRowOrderNumber = toItemRow.CreateCell(B),
                    cellToRowCity = toItemRow.CreateCell(C),
                    cellToWeight = toItemRow.CreateCell(D),
                    cellTotalAmount = toItemRow.CreateCell(E),
                    cellDate = toItemRow.CreateCell(F),
                    cellFirstWeight = toItemRow.CreateCell(G),
                    cellFirstAmount = toItemRow.CreateCell(H),
                    cellFirstAmountB = toItemRow.CreateCell(I),
                    cellOtherAmount = toItemRow.CreateCell(J);

                cellToIndex.SetCellType(CellType.Numeric);
                cellToRowOrderNumber.SetCellType(CellType.String);
                cellDate.SetCellType(CellType.String);
                cellToRowCity.SetCellType(CellType.String);
                cellToWeight.SetCellType(CellType.Numeric);
                cellTotalAmount.SetCellType(CellType.Formula);
                cellFirstWeight.SetCellType(CellType.Numeric);
                cellFirstAmount.SetCellType(CellType.Numeric);
                cellOtherAmount.SetCellType(CellType.Numeric);
                cellFirstAmountB.SetCellType(CellType.Numeric);

                cellToIndex.SetCellValue(index++);
                cellToRowOrderNumber.SetCellValue(ee.OrderNumber);
                if (ee.Flag)
                {
                    if (ee.Weight.Equals(0))
                        cellToRowOrderNumber.CellStyle = toStyleZero;
                    else
                        cellToRowOrderNumber.CellStyle = toStyleGeneric;
                    cellDate.SetCellValue(ee.Date);
                    cellDate.CellStyle = toStyleDate;
                    cellToRowCity.SetCellValue(ee.City);
                    cellToRowCity.CellStyle = toStyleGeneric;
                    cellToWeight.SetCellValue(ee.Weight);
                    cellToWeight.CellStyle = toStyleTotalAmount;
                    Data4Cfg d4c = null;
                    for (int j = 1, jlen = Data.Count; j < jlen; ++j)
                    {
                        Data4Cfg item = Data[j];
                        bool flag = false;
                        string[] citys = item.Key.Split('|');
                        foreach (string c in citys)
                        {
                            if (ee.City != null && ee.City.IndexOf(c) >= 0)
                            {
                                flag = true;
                                break;
                            }
                        }
                        if (flag)
                        {
                            d4c = item;
                            break;
                        }
                    }

                    if (d4c == null)
                        d4c = Data[0];

                    string totalAmount = string.Format("ROUNDUP(IF((D{0}-G{0})<=0,0,ROUNDDOWN((D{0}-G{0}),1)),0)*J{0}+IF((D{0}-G{0})<=0,H{0},I{0})", toSheetRowIndex);
                    cellTotalAmount.SetCellFormula(totalAmount);
                    //if (d4c.Key != "default")
                    //{
                    //    cellTotalAmount.CellStyle = colorStyle;

                    //}
                    //else
                    //{
                    cellTotalAmount.CellStyle = toStyleTotalAmount;
                    //}

                    double firstWeight = Convert.ToDouble(d4c.FirstWeight);
                    cellFirstWeight.SetCellValue(firstWeight);
                    cellFirstWeight.CellStyle = toStyleTotalAmount;
                    double firstAmount = Convert.ToDouble(d4c.FirstAmount);
                    cellFirstAmount.SetCellValue(firstAmount);
                    cellFirstAmount.CellStyle = toStyleTotalAmount;
                    double otherAmount = Convert.ToDouble(d4c.OtherAmount);
                    cellOtherAmount.SetCellValue(otherAmount);
                    cellOtherAmount.CellStyle = toStyleTotalAmount;
                    double firstAmountB = Convert.ToDouble(d4c.FirstAmountB);
                    cellFirstAmountB.SetCellValue(firstAmountB);
                    cellFirstAmountB.CellStyle = toStyleTotalAmount;
                }
                else
                {
                    cellToRowOrderNumber.CellStyle = toStyleNull;
                }
            }

            toSheet.SetColumnWidth(0, 3600);
            toWorkbook.Write(toFile);
        }

        /// <summary>
        /// 电子单
        /// </summary>
        /// <param name="fromFile"></param>
        /// <param name="toFile"></param>
        public void ProcessExcelTemplateB(FileStream fromFile, FileStream toFile)
        {
            XSSFWorkbook workbook = new XSSFWorkbook(fromFile);
            ISheet sheetAt = workbook.GetSheetAt(0);
            ICellStyle style = workbook.CreateCellStyle();
            //style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            //style.FillPattern = FillPattern.SolidForeground;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;

            ICellStyle colorStyle = workbook.CreateCellStyle();
            colorStyle.BorderBottom = BorderStyle.Thin;
            colorStyle.BorderLeft = BorderStyle.Thin;
            colorStyle.BorderRight = BorderStyle.Thin;
            colorStyle.BorderTop = BorderStyle.Thin;
            colorStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
            colorStyle.FillPattern = FillPattern.SolidForeground;

            IDataFormat format = workbook.CreateDataFormat();
            style.DataFormat = format.GetFormat("#,##0.00");


            ICellStyle styleZero = workbook.CreateCellStyle();
            styleZero.BorderBottom = BorderStyle.Thin;
            styleZero.BorderLeft = BorderStyle.Thin;
            styleZero.BorderRight = BorderStyle.Thin;
            styleZero.BorderTop = BorderStyle.Thin;
            styleZero.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
            styleZero.FillPattern = FillPattern.SolidForeground;
            styleZero.DataFormat = format.GetFormat("#,##0.00");

            for (int i = sheetAt.FirstRowNum; i < sheetAt.LastRowNum; ++i)
            {
                IRow row = sheetAt.GetRow(i);
                /*
                0: 空
                1: 标题
                2: 空
                3: 合计
                4: 空
                5: 表头
                */
                if (i == 5)
                {
                    ICellStyle headStyle = row.GetCell(row.LastCellNum - 1).CellStyle;
                    ICell
                        cellTotalAmount = row.CreateCell(L),
                        cellFirstWeight = row.CreateCell(M),
                        cellFirstAmount = row.CreateCell(N),
                        cellFirstAmountB = row.CreateCell(O),
                        cellOtherAmount = row.CreateCell(P);
                    cellTotalAmount.CellStyle = headStyle;
                    cellTotalAmount.SetCellValue("金额");
                    cellFirstWeight.CellStyle = headStyle;
                    cellFirstWeight.SetCellValue("首重重量");
                    cellFirstAmount.CellStyle = headStyle;
                    cellFirstAmount.SetCellValue("小件首重金额");
                    cellFirstAmountB.CellStyle = headStyle;
                    cellFirstAmountB.SetCellValue("大件首重金额");
                    cellOtherAmount.CellStyle = headStyle;
                    cellOtherAmount.SetCellValue("续重金额");
                }
                if (i >= 6)
                {
                    ICell
                        cellOrderNumber = row.GetCell(E) /* E列 */,
                        cellCity = row.GetCell(H) /* H列 */,
                        cellWeight = row.GetCell(K) /* K列 */,
                        cellTotalAmount = row.CreateCell(L),
                        cellFirstWeight = row.CreateCell(M),
                        cellFirstAmount = row.CreateCell(N),
                        cellFirstAmountB = row.CreateCell(O),
                        cellOtherAmount = row.CreateCell(P);
                    string city = cellCity.StringCellValue;
                    double weight = cellWeight.NumericCellValue;

                    Data4Cfg d4c = null;

                    for (int j = 1, jlen = Data.Count; j < jlen; ++j)
                    {
                        Data4Cfg item = Data[j];
                        bool flag = false;
                        string[] citys = item.Key.Split('|');
                        foreach (string c in citys)
                        {
                            if (city.IndexOf(c) >= 0)
                            {
                                flag = true;
                                break;
                            }
                        }
                        if (flag)
                        {
                            d4c = item;
                            break;
                        }
                    }

                    if (d4c == null)
                        d4c = Data[0];
                    cellFirstWeight.SetCellValue(Convert.ToDouble(d4c.FirstWeight));
                    cellFirstWeight.CellStyle = style;
                    cellFirstAmount.SetCellValue(Convert.ToDouble(d4c.FirstAmount));
                    cellFirstAmount.CellStyle = style;
                    cellOtherAmount.SetCellValue(Convert.ToDouble(d4c.OtherAmount));
                    cellOtherAmount.CellStyle = style;
                    cellFirstAmountB.SetCellValue(Convert.ToDouble(d4c.FirstAmountB));
                    cellFirstAmountB.CellStyle = style;
                    //cellTotalAmount.SetCellFormula(string.Format("ROUNDUP(IF((J{0}-L{0})<=0,0,ROUNDDOWN((J{0}-L{0}),1)),0)*N{0}+M{0}", i + 1));
                    cellTotalAmount.SetCellFormula(string.Format("ROUNDUP(IF((K{0}-M{0})<=0,0,ROUNDDOWN((K{0}-M{0}),1)),0)*P{0}+IF((K{0}-M{0})<=0,N{0},O{0})", i + 1));
                    //if (d4c.Key != "default")
                    //{
                    //    cellTotalAmount.CellStyle = colorStyle;

                    //}
                    //else
                    //{
                    cellTotalAmount.CellStyle = style;
                    //}
                    if (weight.Equals(0))
                    {
                        cellOrderNumber.CellStyle = styleZero;
                    }
                }
            }
            workbook.Write(toFile);
        }

        /// <summary>
        /// 菜鸟
        /// </summary>
        /// <param name="fromFile"></param>
        /// <param name="toFile"></param>
        public void ProcessExcelTemplateA(FileStream fromFile, FileStream toFile)
        {
            XSSFWorkbook workbook = new XSSFWorkbook(fromFile);
            ISheet sheetAt = workbook.GetSheetAt(0);
            ICellStyle style = workbook.CreateCellStyle();
            //style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
            //style.FillPattern = FillPattern.SolidForeground;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;

            ICellStyle colorStyle = workbook.CreateCellStyle();
            colorStyle.BorderBottom = BorderStyle.Thin;
            colorStyle.BorderLeft = BorderStyle.Thin;
            colorStyle.BorderRight = BorderStyle.Thin;
            colorStyle.BorderTop = BorderStyle.Thin;
            colorStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
            colorStyle.FillPattern = FillPattern.SolidForeground;


            IDataFormat format = workbook.CreateDataFormat();
            style.DataFormat = format.GetFormat("#,##0.00");


            ICellStyle styleZero = workbook.CreateCellStyle();
            styleZero.BorderBottom = BorderStyle.Thin;
            styleZero.BorderLeft = BorderStyle.Thin;
            styleZero.BorderRight = BorderStyle.Thin;
            styleZero.BorderTop = BorderStyle.Thin;
            styleZero.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index;
            styleZero.FillPattern = FillPattern.SolidForeground;
            styleZero.DataFormat = format.GetFormat("#,##0.00");

            for (int i = sheetAt.FirstRowNum; i < sheetAt.LastRowNum; ++i)
            {
                IRow row = sheetAt.GetRow(i);
                /*
                0: 空
                1: 标题
                2: 空
                3: 合计
                4: 空
                5: 表头
                */
                if (i == 5)
                {
                    ICellStyle headStyle = row.GetCell(row.LastCellNum - 1).CellStyle;
                    ICell
                        cellTotalAmount = row.CreateCell(O),
                        cellFirstWeight = row.CreateCell(P)   /* O */,
                        cellFirstAmount = row.CreateCell(Q)   /* P */,
                        cellFirstAmountB = row.CreateCell(R)  /* Q */,
                        cellOtherAmount = row.CreateCell(S)   /* R */;
                    cellTotalAmount.CellStyle = headStyle;
                    cellTotalAmount.SetCellValue("金额");

                    cellFirstWeight.CellStyle = headStyle;
                    cellFirstWeight.SetCellValue("首重重量");

                    cellFirstAmount.CellStyle = headStyle;
                    cellFirstAmount.SetCellValue("小件首重金额");

                    cellFirstAmountB.CellStyle = headStyle;
                    cellFirstAmountB.SetCellValue("大件首重金额");

                    cellOtherAmount.CellStyle = headStyle;
                    cellOtherAmount.SetCellValue("续重金额");
                }
                if (i >= 6)
                {
                    ICell
                        cellOrderNumber = row.GetCell(E) /* E列 */,
                        cellProvince = row.GetCell(J) /* J列 */,
                        cellCity = row.GetCell(K)  /* K */,
                        cellWeight = row.GetCell(N) /* N列 */,
                        cellTotalAmount = row.CreateCell(O)   /* N */,
                        cellFirstWeight = row.CreateCell(P)   /* O */,
                        cellFirstAmount = row.CreateCell(Q)   /* P */,
                        cellFirstAmountB = row.CreateCell(R)  /* Q */,
                        cellOtherAmount = row.CreateCell(S)   /* R */;
                    string city = cellProvince.StringCellValue + cellCity.StringCellValue;
                    double weight = cellWeight.NumericCellValue;

                    Data4Cfg d4c = null;

                    for (int j = 1, jlen = Data.Count; j < jlen; ++j)
                    {
                        Data4Cfg item = Data[j];
                        bool flag = false;
                        string[] citys = item.Key.Split('|');
                        foreach (string c in citys)
                        {
                            if (city.IndexOf(c) >= 0)
                            {
                                flag = true;
                                break;
                            }
                        }
                        if (flag)
                        {
                            d4c = item;
                            break;
                        }
                    }

                    if (d4c == null)
                        d4c = Data[0];
                    cellFirstWeight.SetCellValue(Convert.ToDouble(d4c.FirstWeight));
                    cellFirstWeight.CellStyle = style;
                    cellFirstAmount.SetCellValue(Convert.ToDouble(d4c.FirstAmount));
                    cellFirstAmount.CellStyle = style;
                    cellFirstAmountB.SetCellValue(Convert.ToDouble(d4c.FirstAmountB));
                    cellFirstAmountB.CellStyle = style;
                    cellOtherAmount.SetCellValue(Convert.ToDouble(d4c.OtherAmount));
                    cellOtherAmount.CellStyle = style;
                    cellTotalAmount.SetCellFormula(string.Format("ROUNDUP(IF((N{0}-P{0})<=0,0,ROUNDDOWN((N{0}-P{0}),1)),0)*S{0}+IF((N{0}-P{0})<=0,Q{0},R{0})", i + 1));
                    //if (d4c.Key != "default")
                    //{
                    //    cellTotalAmount.CellStyle = colorStyle;

                    //}
                    //else
                    //{
                    cellTotalAmount.CellStyle = style;
                    //}
                    if (weight.Equals(0))
                    {
                        cellOrderNumber.CellStyle = styleZero;
                    }
                }
            }
            workbook.Write(toFile);
        }
    }

    public class ExpressEntity
    {
        public string OrderNumber { get; set; }
        public double Date { get; set; }
        public string City { get; set; }
        public double Weight { get; set; }
        public bool Flag { get; set; }
    }
}
