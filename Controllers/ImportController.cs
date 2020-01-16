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
                var rootFolder = Directory.GetCurrentDirectory()+"\\Content";
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
                    ViewBag.Message = companyName + " company not found. So file can't be uploaded, Please contact Admin to add.";
                    return View("Index");
                }

                using (ExcelPackage package = new ExcelPackage(fileLocation))
                {
                    foreach (var workSheet in package.Workbook.Worksheets)
                    {
                        int totalRows = workSheet.Dimension.Rows;
                        string categoryName = workSheet.Name;
                        var category = _context.Categories.FirstOrDefault(m => m.CategoryName == categoryName && m.IsActive == true);

                        long? currentParentLineItemID = null;

                        for (int i = 4; i <= totalRows; i++)
                        {
                            long currentLineItemID = 0;

                            if (category != null && workSheet.Cells[i, 1].Value != null)
                            {
                                var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                var lineItem = _context.LineItems.FirstOrDefault(m => m.CategoryId == category.CategoryId && m.IsActive == true);
                                if (LineItemText.EndsWith(':') || LineItemText.ToLower().Trim().Equals("check"))
                                {
                                    currentParentLineItemID = lineItem.LineItemId;
                                    continue;
                                }
                                if (lineItem.LineItemId > 0)
                                {
                                    currentLineItemID = lineItem.LineItemId;
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
                                        updateLine.ParentLineItemId = currentParentLineItemID;
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
                                    if (LineItemText.EndsWith(':'))
                                    {
                                        currentParentLineItemID = lineItem.LineItemId;
                                    }
                                    newLine.ParentLineItemId = currentParentLineItemID;
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
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Length >= 4)
                                            {
                                                if (yearQuarterText.Contains("Q"))
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
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatements.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.IsActive = true;
                                            finStatement.CreatedDate = DateTime.Now;
                                            _context.FinanceStatements.Add(finStatement);
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

                                            var IS = _context.ISs.Where(m => m.StatementId == currentStatementID && m.LineItemId == currentLineItemID && m.IsActive == true).FirstOrDefault();
                                            if (IS != null)
                                            {
                                                IS.ItemValue = dataValue;
                                                IS.ModifiedDate = DateTime.Now;
                                                _context.ISs.Update(IS);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                IS = new IS();
                                                IS.ItemValue = dataValue;
                                                IS.LineItemId = currentLineItemID;
                                                IS.StatementId = currentStatementID;
                                                IS.IsActive = true;
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
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Length >= 4)
                                            {
                                                if (yearQuarterText.Contains("Q"))
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
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatements.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            _context.FinanceStatements.Add(finStatement);
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
                                                BS.ItemValue = dataValue;
                                                BS.ModifiedDate = DateTime.Now;
                                                _context.BalanceSheets.Update(BS);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                BS = new BalanceSheet();
                                                BS.ItemValue = dataValue;
                                                BS.LineItemId = currentLineItemID;
                                                BS.StatementId = currentStatementID;
                                                BS.IsActive = true;
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
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Length >= 4)
                                            {
                                                //Quarterly texts
                                                if (yearQuarterText.Contains("Q"))
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
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatements.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            finStatement.IsActive = true;
                                            _context.FinanceStatements.Add(finStatement);
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
                                                CF.ItemValue = dataValue;
                                                CF.ModifiedDate = DateTime.Now;
                                                _context.CashFlows.Update(CF);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                CF = new CashFlow();
                                                CF.ItemValue = dataValue;
                                                CF.LineItemId = currentLineItemID;
                                                CF.StatementId = currentStatementID;
                                                CF.IsActive = true;
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
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Length >= 4)
                                            {
                                                //Quarterly texts
                                                if (yearQuarterText.Contains("Q"))
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
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatements.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            finStatement.IsActive = true;
                                            _context.FinanceStatements.Add(finStatement);
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
                                                ISNG.ItemValue = dataValue;
                                                ISNG.ModifiedDate = DateTime.Now;
                                                _context.ISNonGAAPs.Update(ISNG);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                ISNG = new ISNonGAAP();
                                                ISNG.ItemValue = dataValue;
                                                ISNG.LineItemId = currentLineItemID;
                                                ISNG.StatementId = currentStatementID;
                                                ISNG.IsActive = true;
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
                                            if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Length >= 4)
                                            {
                                                if (!string.IsNullOrEmpty(yearQuarterText) && yearQuarterText.Length >= 4)
                                                {
                                                    //Quarterly texts
                                                    if (yearQuarterText.Contains("Q"))
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
                                            }
                                            else
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                        long currentStatementID = 0;
                                        var finStatement = _context.FinanceStatements.FirstOrDefault(m => m.CompanyId == company.CompanyId && m.TransactionDate.Date == dateFormatted.Date && m.Quarter.Contains(yearQuarterText) && m.IsActive == true);
                                        if (finStatement == null)//New Year / Quarter Date to be added
                                        {
                                            finStatement = new FinanceStatement();
                                            finStatement.CompanyId = company.CompanyId;
                                            finStatement.TransactionDate = dateFormatted;
                                            finStatement.Quarter = yearQuarterText;
                                            finStatement.IsHistorical = DateTime.Now.Date.CompareTo(dateFormatted.Date) > 0;
                                            finStatement.CreatedDate = DateTime.Now;
                                            _context.FinanceStatements.Add(finStatement);
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
                                                RD.ItemValue = dataValue;
                                                RD.ModifiedDate = DateTime.Now;
                                                _context.RDs.Update(RD);
                                                _context.SaveChanges();
                                            }
                                            else
                                            {
                                                RD = new RD();
                                                RD.ItemValue = dataValue;
                                                RD.LineItemId = currentLineItemID;
                                                RD.StatementId = currentStatementID;
                                                RD.IsActive = true;
                                                RD.CreatedDate = DateTime.Now;
                                                _context.RDs.Add(RD);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                }
                            }
                            else
                                continue;
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
