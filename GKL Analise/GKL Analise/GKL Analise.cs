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

                var complexity = (double)sheet.Cells[rowNum, (int)ColumnsScan.Сложность].Value;

                var sqFram = (double)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьФрамуги].Value;

                var sqLVs = (double)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьЛевойБоковойВставки].Value;

                var sqRVs = (double)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьПравойБоковойВставки].Value;

                var sqVAS = (double)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовРабочейСтворки].Value;

                var sqVPS = (double)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовПассивнойСтворки].Value;

                var sqVFr = (double)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовФрамуги].Value;

                var sqVLVs = (double)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовЛевойБоковойВставки].Value;

                var sqVRVs = (double)sheet.Cells[rowNum, (int)ColumnsScan.ПлощадьВырезовПравойБоковойВставки].Value;

                var pGab = new Gabaryte(height, width);

                var act = new Stvorka("Активная створка", new Gabaryte(has, was), sqVAS/count);

                Stvorka pas = null;

                Stvorka fr = null;

                Stvorka lv = null;

                Stvorka rv = null;

                if (hps > 0)
                    pas = new Stvorka("Пассивная створка", new Gabaryte(hps, wps), sqVPS/count);

                if (sqLVs > 0)
                    lv = new Stvorka("Левая вставка", sqLVs/count, EGabaryteDirection.Height, height, sqVLVs/count);

                if (sqRVs > 0)
                    rv = new Stvorka("Правая вставка", sqRVs/count, EGabaryteDirection.Height, height, sqVRVs/count);

                if (sqFram > 0)
                    fr = new Stvorka("Фрамуга", sqFram/count, EGabaryteDirection.Width, 
                        width + (lv != null ? lv.Gabaryte.Width : 0) + (rv != null ? rv.Gabaryte.Width : 0), 
                        sqVFr/count);

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

            MessageBox.Show($"Готово\nВсего {dates.AllConstructionCount(EConstructionClass.Product)} изделий");
        }

        private void FillTable(_Worksheet sheet, Dictionary<IConstruction, int> dic, int startCol)
        {
            var row = 3;

            foreach (var d in dic)
            {
                sheet.Cells[row, startCol].Value = d.Key.Gabaryte.Height;
                sheet.Cells[row, startCol+1].Value = d.Key.Gabaryte.Width;
                sheet.Cells[row, startCol+2].Value = d.Value;
                row++;
            }
        }

        private void bFill1_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = (_Worksheet)eApp.ActiveSheet;

            //var dAct = dates
            //    .GetAllConctruction(EConstructionClass.Aktiv)
            //    .OrderByDescending(kvp => kvp.Value)
            //    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value); ;
            //var dPas = dates
            //    .GetAllConctruction(EConstructionClass.Passiv)
            //    .OrderByDescending(kvp => kvp.Value)
            //    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value); ;
            //var dFr = dates
            //    .GetAllConctruction(EConstructionClass.Framuga)
            //    .OrderByDescending(kvp => kvp.Value)
            //    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value); ;
            //var dLV = dates
            //    .GetAllConctruction(EConstructionClass.LVstavka)
            //    .OrderByDescending(kvp => kvp.Value)
            //    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value); ;
            //var dRV = dates
            //    .GetAllConctruction(EConstructionClass.RVstavka)
            //    .OrderByDescending(kvp => kvp.Value)
            //    .ToDictionary(kvp => kvp.Key, kvp => kvp.Value); ;

            //FillTable(sheet, dAct, 1);
            //FillTable(sheet, dPas, 5);
            //FillTable(sheet, dFr, 9);
            //FillTable(sheet, dLV, 13);
            //FillTable(sheet, dRV, 17);

            var d1 = dates.GetAllConctruction(EConstructionClass.Aktiv);
            var d2 = dates.GetAllConctruction(EConstructionClass.Passiv);

            var dic = d1.SumDics(d2);

            var min = dic.GetMinConstruction();
            var max = dic.GetMaxConstruction();
            var mid = dic.GetMidConstruction();
            var mod = dic.GetModaConstruction();
            var m75 = (List<Gabaryte>)dic.Get75Construction();

            sheet.Cells[3, 2].Value = min.Width;
            sheet.Cells[3, 3].Value = min.Height;

            sheet.Cells[4, 2].Value = max.Width;
            sheet.Cells[4, 3].Value = max.Height;

            sheet.Cells[5, 2].Value = mid.Width;
            sheet.Cells[5, 3].Value = mid.Height;

            sheet.Cells[6, 2].Value = mod.Width;
            sheet.Cells[6, 3].Value = mod.Height;

            sheet.Cells[7, 2].Value = $"{m75[0].Width}-{m75[1].Width}";
            sheet.Cells[7, 3].Value = $"{m75[0].Height}-{m75[1].Height}";

            d1 = dates.GetAllConctruction(EConstructionClass.Aktiv);

            min = d1.GetMinConstruction();
            max = d1.GetMaxConstruction();
            mid = d1.GetMidConstruction();
            mod = d1.GetModaConstruction();
            m75 = (List<Gabaryte>)d1.Get75Construction();

            sheet.Cells[11, 2].Value = min.Width;
            sheet.Cells[11, 3].Value = min.Height;

            sheet.Cells[12, 2].Value = max.Width;
            sheet.Cells[12, 3].Value = max.Height;

            sheet.Cells[13, 2].Value = mid.Width;
            sheet.Cells[13, 3].Value = mid.Height;

            sheet.Cells[14, 2].Value = mod.Width;
            sheet.Cells[14, 3].Value = mod.Height;

            sheet.Cells[15, 2].Value = $"{m75[0].Width}-{m75[1].Width}";
            sheet.Cells[15, 3].Value = $"{m75[0].Height}-{m75[1].Height}";

            d1 = dates.GetAllConctruction(EConstructionClass.LVstavka);
            d2 = dates.GetAllConctruction(EConstructionClass.RVstavka);

            dic = d1.SumDics(d2);

            min = dic.GetMinConstruction();
            max = dic.GetMaxConstruction();
            mid = dic.GetMidConstruction();
            mod = dic.GetModaConstruction();
            m75 = (List<Gabaryte>)dic.Get75Construction();

            sheet.Cells[19, 2].Value = min.Width;
            sheet.Cells[19, 3].Value = min.Height;

            sheet.Cells[20, 2].Value = max.Width;
            sheet.Cells[20, 3].Value = max.Height;

            sheet.Cells[21, 2].Value = mid.Width;
            sheet.Cells[21, 3].Value = mid.Height;

            sheet.Cells[22, 2].Value = mod.Width;
            sheet.Cells[22, 3].Value = mod.Height;

            sheet.Cells[23, 2].Value = $"{m75[0].Width}-{m75[1].Width}";
            sheet.Cells[23, 3].Value = $"{m75[0].Height}-{m75[1].Height}";

            dic = dates.GetAllConctruction(EConstructionClass.Framuga);

            min = dic.GetMinConstruction();
            max = dic.GetMaxConstruction();
            mid = dic.GetMidConstruction();
            mod = dic.GetModaConstruction();
            m75 = (List<Gabaryte>)dic.Get75Construction();

            sheet.Cells[27, 2].Value = min.Width;
            sheet.Cells[27, 3].Value = min.Height;

            sheet.Cells[28, 2].Value = max.Width;
            sheet.Cells[28, 3].Value = max.Height;

            sheet.Cells[29, 2].Value = mid.Width;
            sheet.Cells[29, 3].Value = mid.Height;

            sheet.Cells[30, 2].Value = mod.Width;
            sheet.Cells[30, 3].Value = mod.Height;

            sheet.Cells[31, 2].Value = $"{m75[0].Width}-{m75[1].Width}";
            sheet.Cells[31, 3].Value = $"{m75[0].Height}-{m75[1].Height}";

            MessageBox.Show($"Готово");
        }

        private void bFill2_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = (_Worksheet)eApp.ActiveSheet;

            var d1 = dates.GetAllConctruction(EConstructionClass.Aktiv);
            var d2 = dates.GetAllConctruction(EConstructionClass.Passiv);

            var dic = d1.SumDics(d2);

            var diapH = dic.GetDiapHeights();
            var countH = diapH.GetCount();

            var diapW = dic.GetDiapWidth();
            var countW = diapW.GetCount();

            var row = 1;

            sheet.Cells[row, 1].Value = "Створки";
            row++;
            sheet.Cells[row, 1].Value = "Ширина";
            row++;
            sheet.Cells[row, 1].Value = "Интервал, мм";
            sheet.Cells[row, 2].Value = "Количество створок, шт";
            sheet.Cells[row, 3].Value = "Доля, %";

            row++;

            foreach (var dw in diapW) 
            {
                var dol = (double)dw.Value * 100 / countW;

                sheet.Cells[row, 1].Value = dw.Key;
                sheet.Cells[row, 2].Value = dw.Value;
                sheet.Cells[row, 3].Value = dol;

                row++;
            }

            row++;
            row++;

            sheet.Cells[row, 1].Value = "Высота";

            row++;

            sheet.Cells[row, 1].Value = "Интервал, мм";
            sheet.Cells[row, 2].Value = "Количество створок, шт";
            sheet.Cells[row, 3].Value = "Доля, %";

            row++;

            foreach (var dh in diapH)
            {
                var dol = (double)dh.Value * 100 / countH;

                sheet.Cells[row, 1].Value = dh.Key;
                sheet.Cells[row, 2].Value = dh.Value;
                sheet.Cells[row, 3].Value = dol;

                row++;
            }

            row++; 
            row++;
            row++;

            sheet.Cells[row, 1].Value = "Фрамуги";
            row++;
            sheet.Cells[row, 1].Value = "Ширина";
            row++;
            sheet.Cells[row, 1].Value = "Интервал, мм";
            sheet.Cells[row, 2].Value = "Количество створок, шт";
            sheet.Cells[row, 3].Value = "Доля, %";
            row++;

            d1 = dates.GetAllConctruction(EConstructionClass.LVstavka);
            d2 = dates.GetAllConctruction(EConstructionClass.RVstavka);
            var d3 = dates.GetAllConctruction(EConstructionClass.Framuga);

            var d4 = d3.ReversDics();

            dic = d1.SumDics(d2);
            var dic1 = dic.SumDics(d4);

            diapH = dic1.GetDiapHeights();
            countH = diapH.GetCount();

            diapW = dic1.GetDiapWidth();
            countW = diapW.GetCount();

            foreach (var dw in diapW)
            {
                var dol = (double)dw.Value * 100 / countW;

                sheet.Cells[row, 1].Value = dw.Key;
                sheet.Cells[row, 2].Value = dw.Value;
                sheet.Cells[row, 3].Value = dol;

                row++;
            }

            row++;
            row++;

            sheet.Cells[row, 1].Value = "Высота";

            row++;

            sheet.Cells[row, 1].Value = "Интервал, мм";
            sheet.Cells[row, 2].Value = "Количество створок, шт";
            sheet.Cells[row, 3].Value = "Доля, %";

            row++;

            foreach (var dh in diapH)
            {
                var dol = (double)dh.Value * 100 / countH;

                sheet.Cells[row, 1].Value = dh.Key;
                sheet.Cells[row, 2].Value = dh.Value;
                sheet.Cells[row, 3].Value = dol;

                row++;
            } 

            MessageBox.Show($"Готово");
        }
    }
}
