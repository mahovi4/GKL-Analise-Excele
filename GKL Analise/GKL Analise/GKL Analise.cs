using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace GKL_Analise
{
    public partial class Ribbon1
    {
        Excel.Application eApp;

        List<Date> dates = new List<Date>();

        int AllProductsCount = 0;

        Dictionary<int, int> AllHeights = new Dictionary<int, int>();
        Dictionary<int, int> AllWidths = new Dictionary<int, int>();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            eApp = Globals.ThisAddIn.Application;
        }

        private void bScan_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = (_Worksheet)eApp.ActiveSheet;

            var visRange = sheet.UsedRange.Offset[1, 0].SpecialCells(XlCellType.xlCellTypeVisible);

            var wMonth = 0;
            Date date = null;

            foreach (Range row in visRange.Rows)
            {
                var rowNum = row.Row;

                if (rowNum == 1) continue;

                var ch = sheet.Cells[rowNum, (int)ColumnsScan.Год].Value;
                if (ch == null) 
                {
                    dates.Add(date);

                    break;
                }

                var year = (int)sheet.Cells[rowNum, (int)ColumnsScan.Год].Value;

                var month = (int)sheet.Cells[rowNum, (int)ColumnsScan.Месяц].Value;

                var prodName = (string)sheet.Cells[rowNum, (int)ColumnsScan.Продукция].Value;

                var height = (int)sheet.Cells[rowNum, (int)ColumnsScan.ВысотаИзделия].Value;

                var width = (int)sheet.Cells[rowNum, (int)ColumnsScan.ШиринаИзделия].Value;

                var has = (int)sheet.Cells[rowNum, (int)ColumnsScan.ВАС].Value;

                var was = (int)sheet.Cells[rowNum, (int)ColumnsScan.ШАС].Value;

                var hps = (int)sheet.Cells[rowNum, (int)ColumnsScan.ВПС].Value;

                var wps = (int)sheet.Cells[rowNum, (int)ColumnsScan.ШПС].Value;

                var count = (int)sheet.Cells[rowNum, (int)ColumnsScan.Количество].Value;

                var complexity = (int)sheet.Cells[rowNum, (int)ColumnsScan.Сложность].Value;

                var sqFram = (int)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьФрамуги].Value;

                var sqLVs = (int)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьЛевойБоковойВставки].Value;

                var sqRVs = (int)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьПравойБоковойВставки].Value;

                var sqVAS = (int)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовРабочейСтворки].Value;

                var sqVPS = (int)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовПассивнойСтворки].Value;

                var sqVFr = (int)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовФрамуги].Value;

                var sqVLVs = (int)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовЛевойБоковойВставки].Value;

                var sqVRVs = (int)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовПравойБоковойВставки].Value;

                var pGab = new Gabaryte(height, width);

                var act = new Stvorka(new Gabaryte(has, was), sqVAS);

                Stvorka pas = null;

                Framuga fr = null;

                Framuga lv = null;

                Framuga rv = null;

                if (hps > 0)
                    pas = new Stvorka(new Gabaryte(hps, wps), sqVFr);

                if (sqFram > 0)
                    fr = new Framuga(sqFram, sqVFr);

                if (sqLVs > 0)
                    lv = new Framuga(sqLVs, sqVLVs);

                if (sqRVs > 0)
                    rv = new Framuga(sqRVs, sqVRVs);

                if(wMonth == month)
                {
                    date.Products.Add(new Product(prodName, pGab, complexity, act, pas, fr, lv, rv), count);
                }
                else
                {
                    if(date != null)
                        dates.Add(date);

                    date = new Date(year, month);
                    date.Products.Add(new Product(prodName, pGab, complexity, act, pas, fr, lv, rv), count);

                    wMonth = month;
                }
            }

            var b = false;

            foreach (var d in dates)
                foreach (var p in d.Products)
                {
                    AllProductsCount += p.Value;

                    foreach(var h in AllHeights)
                    {
                        b = false;
                        if(h.Key == p.Key.Gabaryte.Height)
                        {
                            AllHeights[h.Key] += p.Value;
                            b = true;
                        }
                    }

                    if (!b) AllHeights.Add(p.Key.Gabaryte.Height, p.Value);
                    
                    foreach(var w in AllWidths)
                    {
                        b = false;
                        if(w.Key == p.Key.Gabaryte.Width)
                        {
                            AllWidths[w.Key] += p.Value;
                            b = true;
                        }
                    }

                    if (!b) AllWidths.Add(p.Key.Gabaryte.Width, p.Value);
                }

            MessageBox.Show($"Готово\nВсего {AllProductsCount} изделий");
        }

        private void bFill1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private Gabaryte GetMin()
        {
            var h = 0;
            var w = 0;

            foreach (var height in AllHeights)
            {
                if (h == 0)
                    h = height.Key;
                else if (h > height.Key)
                    h = height.Key;
            }

            foreach(var width in AllWidths)
            { 

                if (w == 0)
                    w = width.Key;
                else if (w > width.Key)
                    w = width.Key;
            }

            return new Gabaryte(h, w);
        }

        private Gabaryte GetMax()
        {
            var h = 0;
            var w = 0;

            foreach (var d in dates)
                foreach (var p in d.Products)
                {
                    if (h == 0)
                        h = p.Key.Gabaryte.Height;
                    else if (h < p.Key.Gabaryte.Height)
                        h = p.Key.Gabaryte.Height;

                    if (w == 0)
                        w = p.Key.Gabaryte.Width;
                    else if (w < p.Key.Gabaryte.Width)
                        w = p.Key.Gabaryte.Width;
                }

            return new Gabaryte(h, w);
        }

        private Gabaryte GetMid()
        {
            var min = GetMin();
            var max = GetMax();

            return new Gabaryte((min.Height + max.Height)/2, (min.Width + max.Height)/2);
        }

        private Gabaryte GetModa()
        {
            var allWidth = new Dictionary<int, int>();
            var allHeight = new Dictionary<int, int>();

            foreach(var d in dates)
                foreach(var p in d.Products)
                {
                    foreach(var h in allHeight)
                }
        }
    }
}
