using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Proje.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public void Excel(string data)
        {
            string[] mydata = data.Split(';');


            MemoryStream fs = new MemoryStream();
            SpreadsheetDocument xl = SpreadsheetDocument.Create(fs, SpreadsheetDocumentType.Workbook);
            WorkbookPart wbp = xl.AddWorkbookPart();
            WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();
            Workbook wb = new Workbook();

            FileVersion fv = new FileVersion();
            fv.ApplicationName = "Microsoft Office Excel";
            Worksheet ws = new Worksheet();
            SheetData sd = new SheetData();

            Row r = new Row();
            Cell cisim = new Cell();
            cisim.DataType = CellValues.String;
            cisim.CellValue = new CellValue("İSİM");
            r.Append(cisim);


            Cell csoyisim = new Cell();
            csoyisim.DataType = CellValues.String;
            csoyisim.CellValue = new CellValue("SOYİSİM");
            r.Append(csoyisim);

            Cell cadres = new Cell();
            cadres.DataType = CellValues.String;
            cadres.CellValue = new CellValue("ADRES");
            r.Append(cadres);


            Cell cmail = new Cell();
            cmail.DataType = CellValues.String;
            cmail.CellValue = new CellValue("MAİL");
            r.Append(cmail);

            sd.Append(r);

            for (int i = 0; i < mydata.Length; i++)
            {
                string[] cell = mydata[i].Split(',');
                Row r1 = new Row();
                for (int j = 0; j < cell.Length; j++)
                {

                    Cell c1 = new Cell();
                    c1.DataType = CellValues.String;
                    c1.CellValue = new CellValue(cell[j]);
                    r1.Append(c1);
                }

                sd.Append(r1);
            }

            ws.Append(sd);

            wsp.Worksheet = ws;
            wsp.Worksheet.Save();
            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet();
            sheet.Name = "rapor";
            sheet.SheetId = 1;
            sheet.Id = wbp.GetIdOfPart(wsp);
            sheets.Append(sheet);
            wb.Append(fv);
            wb.Append(sheets);

            xl.WorkbookPart.Workbook = wb;
            xl.WorkbookPart.Workbook.Save();
            xl.Close();

            Response.Clear();
            byte[] dt = fs.ToArray();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", string.Format("attachment; filename=Rapor.xlsx"));
            Response.BinaryWrite(dt);
            Response.End();
        }

        public ActionResult OpenExcel(string data)
        {
            Excel(data);
            return new EmptyResult();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}