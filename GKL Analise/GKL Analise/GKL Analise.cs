using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

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
            var date = new Date(0, 0);

            foreach (Range row in visRange.Rows)
            {
                var rowNum = row.Row;

                if (rowNum == 1) continue;

                if(sheet.Cells[rowNum, (int)Columns.Год] == "") return;

                if (!int.TryParse(sheet.Cells[rowNum, (int)Columns.Год].Value, out int year))
                    throw new Exception();

                if (!int.TryParse(sheet.Cells[rowNum, (int)Columns.Месяц].Value, out int month))
                    throw new Exception();

                var prodName = (string)sheet.Cells[rowNum, (int)Columns.Продукция].Value;

                if (!int.TryParse(sheet.Cells[rowNum, (int)Columns.ВысотаИзделия].Value, out int height))
                    throw new Exception();

                if (!int.TryParse(sheet.Cells[rowNum, (int)Columns.ШиринаИзделия].Value, out int width))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ВАС].Value, out int has))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ШАС].Value, out int was))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ВПС].Value, out int hps))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ШПС].Value, out int wps))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.Количество].Value, out int count))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.Сложность].Value, out int complexity))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ПлощадьФрамуги].Value, out int sqFram))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ПлощадьЛевойБоковойВставки].Value, out int sqLVs))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ПлощадьПравойБоковойВставки].Value, out int sqRVs))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ПлощадьВырезовРабочейСтворки].Value, out int sqVAS))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ПлощадьВырезовПассивнойСтворки].Value, out int sqVPS))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ПлощадьВырезовФрамуги].Value, out int sqVFr))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ПлощадьВырезовЛевойБоковойВставки].Value, out int sqVLVs))
                    throw new Exception();

                if(!int.TryParse(sheet.Cells[rowNum, (int)Columns.ПлощадьВырезовПравойБоковойВставки].Value, out int sqVRVs))
                    throw new Exception();

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
                    
                }
            }
        }
    }
}
