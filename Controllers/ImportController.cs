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

            string fileExtension = Path.GetExtension(file.FileName);

            if (fileExtension == ".xls" || fileExtension == ".xlsm" || fileExtension == ".xlsx")
            {
                var rootFolder = @"C:\Content";
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
                    ViewBag.Message = companyName + " company not found. So file can't be uploaded, Please contact admin to add.";
                    return View("Index");
                }

                using (ExcelPackage package = new ExcelPackage(fileLocation))
                {
                    foreach (var workSheet in package.Workbook.Worksheets)
                    {
                        //ExcelWorksheet workSheet = package.Workbook.Worksheets["Table1"];
                        ////var workSheet = package.Workbook.Worksheets.First();
                        int totalRows = workSheet.Dimension.Rows;
                        string categoryName = workSheet.Name;
                        var category = _context.Categories.FirstOrDefault(m => m.CategoryName == categoryName && m.IsActive == true);

                        for (int i = 4; i <= totalRows; i++)
                        {
                            long currentLineItemID = 0;
                            if (category != null)
                            {
                                var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                var lineItemId = _context.LineItems.FirstOrDefault(m => m.CategoryId == category.CategoryId && m.IsActive == true).LineItemId;
                                if (lineItemId > 0)
                                {
                                    currentLineItemID = lineItemId;
                                    LineItem updateLine = _context.LineItems.FirstOrDefault(m => m.LineItemText == LineItemText && m.CategoryId == category.CategoryId && m.IsActive == true);
                                    if (updateLine != null)
                                    {
                                        updateLine.LineItemText = LineItemText;
                                        updateLine.ModifiedDate = DateTime.Now;
                                        _context.LineItems.Update(updateLine);
                                        _context.SaveChanges();
                                    }
                                    else
                                    {
                                        updateLine = new LineItem();
                                        updateLine.LineItemText = LineItemText;
                                        updateLine.ParentLineItemId = null;
                                        updateLine.IsActive = true;
                                        updateLine.CategoryId = category.CategoryId;
                                        updateLine.CreatedDate = DateTime.Now;
                                        _context.LineItems.Add(updateLine);
                                        _context.SaveChanges();
                                        currentLineItemID = updateLine.LineItemId;
                                    }
                                }
                                else
                                {
                                    LineItem newLine = new LineItem();
                                    newLine.LineItemText = LineItemText;
                                    newLine.ParentLineItemId = null;
                                    newLine.IsActive = true;
                                    newLine.CategoryId = category.CategoryId;
                                    newLine.CreatedDate = DateTime.Now;
                                    _context.LineItems.Add(newLine);
                                    _context.SaveChanges();
                                    currentLineItemID = newLine.LineItemId;
                                }

                                //Different worksheets of excel data to be merged
                                //IS Insertion/Updation
                                if (category.CategoryName.Equals("IS"))
                                {
                                    int totalCols = workSheet.Dimension.Columns;
                                    for (int j = 3; j <= totalCols; j++)
                                    {
                                        string yearQuarterText = string.Empty;
                                        string dateText = string.Empty;
                                        DateTime dateFormatted = DateTime.Now;
                                        if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                        {
                                            yearQuarterText = workSheet.Cells[1, j].Value.ToString();//Years or Quarters fetch
                                            //Quarterly texts
                                            if(!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Contains("Q"))
                                            {
                                                yearQuarterText = yearQuarterText.Substring(2, 2);
                                                yearQuarterText = "20" + yearQuarterText;
                                            }
                                            //Else Yearly texts
                                            dateText = workSheet.Cells[2, j].Text.ToString();//Dates fetch
                                            dateText = dateText + "-"+ yearQuarterText.Substring(0, 4);
                                            dateFormatted = DateTime.ParseExact(dateText, "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatement.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            _context.FinanceStatement.Add(finStatement);
                                            _context.SaveChanges();
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        else
                                        {
                                            currentStatementID = finStatement.StatementId;
                                        }

                                        if (workSheet.Cells[i, j].Value != null)
                                        {
                                            var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch

                                            var IS = _context.ISs.FirstOrDefault(m => m.StatementId == currentStatementID && m.LineItemId == currentLineItemID && m.IsActive == true);
                                            if (IS != null)
                                            {
                                                IS.ItemValue = long.Parse(dataValue);
                                                IS.ModifiedDate = DateTime.Now;
                                                _context.ISs.Update(IS);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                IS = new IS();
                                                IS.ItemValue = long.Parse(dataValue);
                                                IS.LineItemId = currentLineItemID;
                                                IS.StatementId = currentStatementID;
                                                IS.CreatedDate = DateTime.Now;
                                                _context.ISs.Add(IS);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                }

                                //BS Insertion/Updation
                                if (category.CategoryName.Equals("BS"))
                                {
                                    int totalCols = workSheet.Dimension.Columns;
                                    for (int j = 3; j <= totalCols; j++)
                                    {
                                        string yearQuarterText = string.Empty;
                                        string dateText = string.Empty;
                                        DateTime dateFormatted = DateTime.Now;
                                        if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                        {
                                            yearQuarterText = workSheet.Cells[1, j].Value.ToString();//Years or Quarters fetch
                                            //Quarterly texts
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Contains("Q"))
                                            {
                                                yearQuarterText = yearQuarterText.Substring(2, 2);
                                                yearQuarterText = "20" + yearQuarterText;
                                            }
                                            //Else Yearly texts
                                            dateText = workSheet.Cells[2, j].Text.ToString();//Dates fetch
                                            dateText = dateText + "-" + yearQuarterText.Substring(0, 4);
                                            dateFormatted = DateTime.ParseExact(dateText, "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatement.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            _context.FinanceStatement.Add(finStatement);
                                            _context.SaveChanges();
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        else
                                        {
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        if (workSheet.Cells[i, j].Value != null)
                                        {
                                            var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch

                                            var BS = _context.BalanceSheets.FirstOrDefault(m => m.StatementId == currentStatementID && m.LineItemId == currentLineItemID && m.IsActive == true);
                                            if (BS != null)
                                            {
                                                BS.ItemValue = long.Parse(dataValue);
                                                BS.ModifiedDate = DateTime.Now;
                                                _context.BalanceSheets.Update(BS);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                BS = new BalanceSheet();
                                                BS.ItemValue = long.Parse(dataValue);
                                                BS.LineItemId = currentLineItemID;
                                                BS.StatementId = currentStatementID;
                                                BS.CreatedDate = DateTime.Now;
                                                _context.BalanceSheets.Add(BS);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                }

                                //CF Insertion/Updation
                                if (category.CategoryName.Equals("CF"))
                                {
                                    int totalCols = workSheet.Dimension.Columns;
                                    for (int j = 3; j <= totalCols; j++)
                                    {
                                        string yearQuarterText = string.Empty;
                                        string dateText = string.Empty;
                                        DateTime dateFormatted = DateTime.Now;
                                        if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                        {
                                            yearQuarterText = workSheet.Cells[1, j].Value.ToString();//Years or Quarters fetch
                                            //Quarterly texts
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Contains("Q"))
                                            {
                                                yearQuarterText = yearQuarterText.Substring(2, 2);
                                                yearQuarterText = "20" + yearQuarterText;
                                            }
                                            //Else Yearly texts
                                            dateText = workSheet.Cells[2, j].Text.ToString();//Dates fetch
                                            dateText = dateText + "-" + yearQuarterText.Substring(0, 4);
                                            dateFormatted = DateTime.ParseExact(dateText, "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatement.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            _context.FinanceStatement.Add(finStatement);
                                            _context.SaveChanges();
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        else
                                        {
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        if (workSheet.Cells[i, j].Value != null)
                                        {
                                            var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch

                                            var CF = _context.CashFlows.FirstOrDefault(m => m.StatementId == currentStatementID && m.LineItemId == currentLineItemID && m.IsActive == true);
                                            if (CF != null)
                                            {
                                                CF.ItemValue = long.Parse(dataValue);
                                                CF.ModifiedDate = DateTime.Now;
                                                _context.CashFlows.Update(CF);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                CF = new CashFlow();
                                                CF.ItemValue = long.Parse(dataValue);
                                                CF.LineItemId = currentLineItemID;
                                                CF.StatementId = currentStatementID;
                                                CF.CreatedDate = DateTime.Now;
                                                _context.CashFlows.Add(CF);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                }

                                //ISNG Insertion/Updation
                                if (category.CategoryName.Equals("IS_NG"))
                                {
                                    int totalCols = workSheet.Dimension.Columns;
                                    for (int j = 3; j <= totalCols; j++)
                                    {
                                        string yearQuarterText = string.Empty;
                                        string dateText = string.Empty;
                                        DateTime dateFormatted = DateTime.Now;
                                        if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                        {
                                            yearQuarterText = workSheet.Cells[1, j].Value.ToString();//Years or Quarters fetch
                                            //Quarterly texts
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Contains("Q"))
                                            {
                                                yearQuarterText = yearQuarterText.Substring(2, 2);
                                                yearQuarterText = "20" + yearQuarterText;
                                            }
                                            //Else Yearly texts
                                            dateText = workSheet.Cells[2, j].Text.ToString();//Dates fetch
                                            dateText = dateText + "-" + yearQuarterText.Substring(0, 4);
                                            dateFormatted = DateTime.ParseExact(dateText, "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatement.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            _context.FinanceStatement.Add(finStatement);
                                            _context.SaveChanges();
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        else
                                        {
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        if (workSheet.Cells[i, j].Value != null)
                                        {
                                            var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch

                                            var ISNG = _context.ISNonGAAPs.FirstOrDefault(m => m.StatementId == currentStatementID && m.LineItemId == currentLineItemID && m.IsActive == true);
                                            if (ISNG != null)
                                            {
                                                ISNG.ItemValue = long.Parse(dataValue);
                                                ISNG.ModifiedDate = DateTime.Now;
                                                _context.ISNonGAAPs.Update(ISNG);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                ISNG = new ISNonGAAP();
                                                ISNG.ItemValue = long.Parse(dataValue);
                                                ISNG.LineItemId = currentLineItemID;
                                                ISNG.StatementId = currentStatementID;
                                                ISNG.CreatedDate = DateTime.Now;
                                                _context.ISNonGAAPs.Add(ISNG);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                }

                                //RD Insertion/Updation
                                if (category.CategoryName.Equals("RD"))
                                {
                                    int totalCols = workSheet.Dimension.Columns;
                                    for (int j = 3; j <= totalCols; j++)
                                    {
                                        string yearQuarterText = string.Empty;
                                        string dateText = string.Empty;
                                        DateTime dateFormatted = DateTime.Now;
                                        if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                        {
                                            yearQuarterText = workSheet.Cells[1, j].Value.ToString();//Years or Quarters fetch
                                            //Quarterly texts
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Contains("Q"))
                                            {
                                                yearQuarterText = yearQuarterText.Substring(2, 2);
                                                yearQuarterText = "20" + yearQuarterText;
                                            }
                                            //Else Yearly texts
                                            dateText = workSheet.Cells[2, j].Text.ToString();//Dates fetch
                                            dateText = dateText + "-" + yearQuarterText.Substring(0, 4);
                                            dateFormatted = DateTime.ParseExact(dateText, "dd-MMM-yyyy", CultureInfo.InvariantCulture);
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatement.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            _context.FinanceStatement.Add(finStatement);
                                            _context.SaveChanges();
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        else
                                        {
                                            currentStatementID = finStatement.StatementId;
                                        }
                                        if (workSheet.Cells[i, j].Value != null)
                                        {
                                            var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch

                                            var RD = _context.RDs.FirstOrDefault(m => m.StatementId == currentStatementID && m.LineItemId == currentLineItemID && m.IsActive == true);
                                            if (RD != null)
                                            {
                                                RD.ItemValue = long.Parse(dataValue);
                                                RD.ModifiedDate = DateTime.Now;
                                                _context.RDs.Update(RD);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                RD = new RD();
                                                RD.ItemValue = long.Parse(dataValue);
                                                RD.LineItemId = currentLineItemID;
                                                RD.StatementId = currentStatementID;
                                                RD.CreatedDate = DateTime.Now;
                                                _context.RDs.Add(RD);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
                ViewBag.Message = "Data Imported Successfully.";
                return View("Index");
            }
            ViewBag.Message = "Other than excel file not allowed.";
            return View("Index");
        }
    }

}
