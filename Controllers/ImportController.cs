using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Web;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using FinResearch.Models;
using System.Globalization;
using Newtonsoft.Json;
using System.Text;

namespace FinResearch.Controllers
{
    public class ImportController : Controller
    {
        private readonly FinResearchContext _context;
        public ImportController(FinResearchContext context)
        {
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ImportFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ViewBag.Message = "No File Found Selected.";
                return View("Index");
            }
            try
            {
                string fileExtension = Path.GetExtension(file.FileName);

                if (fileExtension == ".xls" || fileExtension == ".xlsm" || fileExtension == ".xlsx")
                {
                    var rootFolder = Directory.GetCurrentDirectory() + "\\Content";
                    var fileName = file.FileName;
                    var filePath = Path.Combine(rootFolder, fileName);
                    var fileLocation = new FileInfo(filePath);

                    if (file.Length <= 0)
                    {
                        ViewBag.Message = "Any file not selected.";
                        return View("Index");
                    }
                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(fileStream);
                    }

                    string companyName = file.FileName.Substring(0, file.FileName.Length > 6 ? 6 : file.FileName.Length);
                    string companyCode = string.Empty;
                    if (companyName.Contains(" "))
                    {
                        companyCode = companyName.Split(' ')[0].ToString();
                    }
                    else
                    {
                        companyCode = file.FileName.Substring(0, file.FileName.Length > 3 ? 3 : file.FileName.Length);
                    }

                    var company = _context.Companies.FirstOrDefault(m => m.CompanyCode == companyCode && m.IsActive == true);
                    if (company == null)
                    {
                        ViewBag.Message = companyName + " company not found. So file can't be uploaded, Please contact Admin to add new.";
                        return View("Index");
                    }

                    using (ExcelPackage package = new ExcelPackage(fileLocation))
                    {
                        foreach (var workSheet in package.Workbook.Worksheets)
                        {
                            int totalRows = workSheet.Dimension.Rows;
                            string categoryName = workSheet.Name;
                            var category = _context.Categories.FirstOrDefault(m => m.CategoryName == categoryName && m.IsActive == true);

                            /// Addition of each Line Items as per each excel sheets
                            //Different worksheets of excel datafile to be merged
                            //IS Insertion/Updation
                            if (category != null && category.CategoryName.Equals("IS"))
                            {
                                var JSONString = new StringBuilder();
                                JSONString.Append("[");
                                for (int i = 4; i <= totalRows; i++)
                                {
                                    if (workSheet.Cells[i, 1].Value != null)
                                    {
                                        var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                        if (LineItemText.ToLower().Trim().Equals("check"))
                                        {
                                            continue;
                                        }
                                        if (workSheet.Cells[i, 2].Value != null)
                                        {
                                            JSONString.Append("\"Config\":" + "\"" + workSheet.Cells[i, 2].Value + "\",");
                                        }
                                        else if (LineItemText.EndsWith(':'))
                                        {
                                            JSONString.Append("{");                                                                               
                                            JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\"");
                                            JSONString.Append("}");
                                            continue;
                                        }

                                        JSONString.Append("{");
                                        JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\",");

                                        int totalCols = workSheet.Dimension.Columns;
                                        for (int j = 3; j <= totalCols; j++)
                                        {
                                            if (j < totalCols - 1)
                                            {                                              
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\",");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\",");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\",");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                            else if (j == totalCols - 1)
                                            {                                               
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\"");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\"");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\"");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }

                                                }
                                            }                                           
                                        }
                                        if (i == totalRows - 1)
                                        {
                                            JSONString.Append("}");
                                        }
                                        else
                                        {
                                            JSONString.Append("},");
                                        }
                                    }
                                }
                                JSONString.Append("]");
                                var IS = new ISs();
                                IS.CompanyId  = company.CompanyId;
                                IS.CategoryId = category.CategoryId;
                                IS.Payload = JSONString.ToString();
                                IS.IsActive = true;
                                IS.CreatedDate = DateTime.Now;
                                _context.ISs.Add(IS);
                                _context.SaveChanges();
                            }

                            //BS Insertion/Updation
                            if (category != null && category.CategoryName.Equals("BS"))
                            {                            
                                var JSONString = new StringBuilder();
                                JSONString.Append("[");
                                for (int i = 4; i <= totalRows; i++)
                                {
                                    if (workSheet.Cells[i, 1].Value != null)
                                    {
                                        var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                        if (LineItemText.ToLower().Trim().Equals("check"))
                                        {
                                            continue;
                                        }
                                        if (workSheet.Cells[i, 2].Value != null)
                                        {
                                            JSONString.Append("\"Config\":" + "\"" + workSheet.Cells[i, 2].Value + "\",");
                                        }
                                        else if (LineItemText.EndsWith(':'))
                                        {
                                            JSONString.Append("{");
                                            JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\"");
                                            JSONString.Append("}");
                                            continue;
                                        }

                                        JSONString.Append("{");
                                        JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\",");

                                        int totalCols = workSheet.Dimension.Columns;
                                        for (int j = 2; j <= totalCols; j++)
                                        {
                                            if (j < totalCols - 1)
                                            {                                              
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\",");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\",");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\",");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                            else if (j == totalCols - 1)
                                            {
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\"");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\"");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\"");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                        }
                                        if (i == totalRows - 1)
                                        {
                                            JSONString.Append("}");
                                        }
                                        else
                                        {
                                            JSONString.Append("},");
                                        }
                                    }
                                }
                                JSONString.Append("]");
                                var BS = new BalanceSheets();
                                BS.CompanyId = company.CompanyId;
                                BS.CategoryId = category.CategoryId;
                                BS.Payload = JSONString.ToString();
                                BS.IsActive = true;
                                BS.CreatedDate = DateTime.Now;
                                _context.BalanceSheets.Add(BS);
                                _context.SaveChanges();
                            }

                            //CF Insertion/Updation
                            if (category != null && category.CategoryName.Equals("CF"))
                            {
                                var JSONString = new StringBuilder();
                                JSONString.Append("[");
                                for (int i = 4; i <= totalRows; i++)
                                {
                                    if (workSheet.Cells[i, 1].Value != null)
                                    {
                                        var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                        if (LineItemText.ToLower().Trim().Equals("check"))
                                        {
                                            continue;
                                        }
                                        if (workSheet.Cells[i, 2].Value != null)
                                        {
                                            JSONString.Append("\"Config\":" + "\"" + workSheet.Cells[i, 2].Value + "\",");
                                        }
                                        else if (LineItemText.EndsWith(':'))
                                        {
                                            JSONString.Append("{");
                                            JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\"");
                                            JSONString.Append("}");
                                            continue;
                                        }

                                        JSONString.Append("{");
                                        JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\",");

                                        int totalCols = workSheet.Dimension.Columns;
                                        for (int j = 2; j <= totalCols; j++)
                                        {
                                            if (j < totalCols - 1)
                                            {
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\",");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\",");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\",");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                            else if (j == totalCols - 1)
                                            {
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\"");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\"");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\"");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                        }
                                        if (i == totalRows - 1)
                                        {
                                            JSONString.Append("}");
                                        }
                                        else
                                        {
                                            JSONString.Append("},");
                                        }
                                    }
                                }
                                JSONString.Append("]");

                                var CF = new CashFlows();
                                CF.CompanyId = company.CompanyId;
                                CF.CategoryId = category.CategoryId;
                                CF.Payload = JSONString.ToString();
                                CF.IsActive = true;
                                CF.CreatedDate = DateTime.Now;
                                _context.CashFlows.Add(CF);
                                _context.SaveChanges();
                            }

                            //ISNG Insertion/Updation
                            if (category != null && category.CategoryName.Equals("IS_NG"))
                            {
                                var JSONString = new StringBuilder();
                                JSONString.Append("[");
                                for (int i = 4; i <= totalRows; i++)
                                {
                                    if (workSheet.Cells[i, 1].Value != null)
                                    {
                                        var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                        if (LineItemText.ToLower().Trim().Equals("check"))
                                        {
                                            continue;
                                        }
                                        if (workSheet.Cells[i, 2].Value != null)
                                        {
                                            JSONString.Append("\"Config\":" + "\"" + workSheet.Cells[i, 2].Value + "\",");
                                        }
                                        else if (LineItemText.EndsWith(':'))
                                        {
                                            JSONString.Append("{");
                                            JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\"");
                                            JSONString.Append("}");
                                            continue;
                                        }

                                        int totalCols = workSheet.Dimension.Columns;
                                        for (int j = 2; j <= totalCols; j++)
                                        {
                                            if (j < totalCols - 1)
                                            {
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\",");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\",");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\",");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                            else if (j == totalCols - 1)
                                            {                                             
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\"");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\"");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\"");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                        }
                                        if (i == totalRows - 1)
                                        {
                                            JSONString.Append("}");
                                        }
                                        else
                                        {
                                            JSONString.Append("},");
                                        }
                                    }
                                }
                                JSONString.Append("]");
                                var ISNG = new ISNonGAAPs();
                                ISNG.CompanyId = company.CompanyId;
                                ISNG.CategoryId = category.CategoryId;
                                ISNG.Payload = JSONString.ToString();
                                ISNG.IsActive = true;
                                ISNG.CreatedDate = DateTime.Now;
                                _context.ISNonGAAPs.Add(ISNG);
                                _context.SaveChanges();
                            }

                            //RD Insertion/Updation
                            if (category != null && category.CategoryName.Equals("RD"))
                            {
                                var JSONString = new StringBuilder();
                                JSONString.Append("[");
                                for (int i = 4; i <= totalRows; i++)
                                {
                                    if (workSheet.Cells[i, 1].Value != null)
                                    {
                                        var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                        if (LineItemText.ToLower().Trim().Equals("check"))
                                        {
                                            continue;
                                        }
                                        if (workSheet.Cells[i, 2].Value != null)
                                        {
                                            JSONString.Append("\"Config\":" + "\"" + workSheet.Cells[i, 2].Value + "\",");
                                        }
                                        else if (LineItemText.EndsWith(':'))
                                        {
                                            JSONString.Append("{");
                                            JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\"");
                                            JSONString.Append("}");
                                            continue;
                                        }

                                        JSONString.Append("{");
                                        JSONString.Append("\"LineItem\":" + "\"" + LineItemText + "\",");

                                        int totalCols = workSheet.Dimension.Columns;
                                        for (int j = 2; j <= totalCols; j++)
                                        {
                                            if (j < totalCols - 1)
                                            {
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\",");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\",");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\",");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                            else if (j == totalCols - 1)
                                            {
                                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                                {
                                                    //JSONString.Append("\"YearQuarter\":" + "\"" + workSheet.Cells[1, j].Value + "\"");
                                                    //JSONString.Append("\"Month\":" + "\"" + workSheet.Cells[2, j].Value + "\"");
                                                    if (workSheet.Cells[i, j].Value != null)
                                                    {
                                                        var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"" + dataValue + "\"");
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        JSONString.Append("\"" + workSheet.Cells[1, j].Value + " " + workSheet.Cells[2, j].Text + "\":" + "\"0\"");
                                                    }
                                                }
                                            }
                                        }
                                        if (i == totalRows - 1)
                                        {
                                            JSONString.Append("}");
                                        }
                                        else
                                        {
                                            JSONString.Append("},");
                                        }
                                    }
                                }
                                JSONString.Append("]");
                                var RD = new RDs();
                                RD.CompanyId = company.CompanyId;
                                RD.CategoryId = category.CategoryId;
                                RD.Payload = JSONString.ToString();
                                RD.IsActive = true;
                                RD.CreatedDate = DateTime.Now;
                                _context.RDs.Add(RD);
                                _context.SaveChanges();
                            }

                        }
                    }
                    ViewBag.Message = "Data Imported Successfully.";
                    return View("Index");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            ViewBag.Message = "Other than excel file not allowed.";
            return View("Index");
        }
    }

}
