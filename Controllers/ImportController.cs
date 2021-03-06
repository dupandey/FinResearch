﻿using System;
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
using Newtonsoft.Json.Linq;
using System.Security.Claims;

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
                                //var JSONString = new StringBuilder();
                                JArray jArray = GetExcelArray(workSheet, totalRows);
                                JArray jthemeArray = GetExcelThemeArray(workSheet, totalRows);
                                var IS = new ISs();
                                IS.CompanyId = company.CompanyId;
                                IS.CategoryId = category.CategoryId;
                                IS.Payload = jArray.ToString();
                                IS.FileName = fileName;
                                IS.FileVersion = 0;
                                IS.IsAdmin = true;
                                IS.UserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
                                IS.IsActive = true;
                                IS.CreatedDate = DateTime.Now;
                                _context.ISs.Add(IS);

                                var themeis = new ThemeMasterIs();
                                themeis.CompanyId = company.CompanyId;
                                themeis.CategoryId = category.CategoryId;
                                themeis.Payload = jthemeArray.ToString();
                                themeis.IsActive = true;
                                themeis.CreatedDate = DateTime.Now;
                                _context.ThemeMasterIss.Add(themeis);

                                _context.SaveChanges();

                            }

                            //BS Insertion/Updation
                            if (category != null && category.CategoryName.Equals("BS"))
                            {
                                JArray jArray = GetExcelArray(workSheet, totalRows);
                                var BS = new BalanceSheets();
                                BS.CompanyId = company.CompanyId;
                                BS.CategoryId = category.CategoryId;
                                BS.Payload = jArray.ToString();
                                BS.FileName = fileName;
                                BS.FileVersion = 0;
                                BS.IsAdmin = true;
                                BS.UserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
                                BS.IsActive = true;
                                BS.CreatedDate = DateTime.Now;
                                _context.BalanceSheets.Add(BS);
                                _context.SaveChanges();
                            }

                            //CF Insertion/Updation
                            if (category != null && category.CategoryName.Equals("CF"))
                            {

                                JArray jArray = GetExcelArray(workSheet, totalRows);
                                var CF = new CashFlows();
                                CF.CompanyId = company.CompanyId;
                                CF.CategoryId = category.CategoryId;
                                CF.Payload = jArray.ToString();
                                CF.FileName = fileName;
                                CF.FileVersion = 0;
                                CF.IsAdmin = true;
                                CF.UserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
                                CF.IsActive = true;
                                CF.CreatedDate = DateTime.Now;
                                _context.CashFlows.Add(CF);
                                _context.SaveChanges();
                            }

                            //ISNG Insertion/Updation
                            if (category != null && category.CategoryName.Equals("IS_NG"))
                            {
                                var JSONString = new StringBuilder();
                                JArray jArray = new JArray();
                                for (int i = 4; i <= totalRows; i++)
                                {
                                    JObject objNG = new JObject();
                                    if (workSheet.Cells[i, 1].Value != null)
                                    {
                                        var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                        if (LineItemText.ToLower().Trim().Equals("check"))
                                        {
                                            continue;
                                        }
                                        objNG.Add("Config", workSheet.Cells[i, 2].Value == null ? " " : workSheet.Cells[i, 2].Value.ToString());
                                        objNG.Add("LineItem", LineItemText);

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

                                                        objNG.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, dataValue);
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        objNG.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, 0);
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
                                                        objNG.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, dataValue);
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        objNG.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, 0);
                                                    }
                                                }
                                            }
                                        }

                                    }
                                    if (objNG.HasValues)
                                        jArray.Add(objNG);
                                }
                                var ISNG = new ISNonGAAPs();
                                ISNG.CompanyId = company.CompanyId;
                                ISNG.CategoryId = category.CategoryId;
                                ISNG.Payload = jArray.ToString();
                                ISNG.FileName = fileName;
                                ISNG.FileVersion = 0;
                                ISNG.IsAdmin = true;
                                ISNG.UserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
                                ISNG.IsActive = true;
                                ISNG.CreatedDate = DateTime.Now;
                                _context.ISNonGAAPs.Add(ISNG);
                                _context.SaveChanges();
                            }

                            //RD Insertion/Updation
                            if (category != null && category.CategoryName.Equals("RD"))
                            {
                                JArray jArray = new JArray();
                                for (int i = 4; i <= totalRows; i++)
                                {
                                    JObject objRD = new JObject();
                                    if (workSheet.Cells[i, 1].Value != null)
                                    {
                                        var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                                        if (LineItemText.ToLower().Trim().Equals("check"))
                                        {
                                            continue;
                                        }
                                        objRD.Add("Config", workSheet.Cells[i, 2].Value == null ? " " : workSheet.Cells[i, 2].Value.ToString());
                                        objRD.Add("LineItem", LineItemText);

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

                                                        objRD.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, dataValue);
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        objRD.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, 0);
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
                                                        objRD.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, dataValue);
                                                    }
                                                    else
                                                    {
                                                        //Blank 0 data value will be added
                                                        objRD.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, 0);
                                                    }
                                                }
                                            }
                                        }

                                    }
                                    if (objRD.HasValues)
                                        jArray.Add(objRD);
                                }

                                var RD = new RDs();
                                RD.CompanyId = company.CompanyId;
                                RD.CategoryId = category.CategoryId;
                                RD.Payload = jArray.ToString();
                                RD.FileName = fileName;
                                RD.FileVersion = 0;
                                RD.IsAdmin = true;
                                RD.UserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
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

        private static JArray GetExcelArray(ExcelWorksheet workSheet, int totalRows)
        {
            JArray jArray = new JArray();
            for (int i = 4; i <= totalRows; i++)
            {
                JObject jObject = new JObject();
                if (workSheet.Cells[i, 1].Value != null)
                {
                    var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                    if (LineItemText.ToLower().Trim().Equals("check"))
                    {
                        continue;
                    }
                    jObject.Add("Config", workSheet.Cells[i, 2].Value == null ? " " : workSheet.Cells[i, 2].Value.ToString());
                    jObject.Add("LineItem", LineItemText);

                    int totalCols = workSheet.Dimension.Columns;
                    for (int j = 2; j <= totalCols; j++)
                    {
                        if (j < totalCols - 1)
                        {
                            if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                            {
                                if (workSheet.Cells[i, j].Value != null)
                                {
                                    var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                    decimal result = 0;
                                    bool canConvert = decimal.TryParse(dataValue, out result);
                                    if (canConvert == true)
                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, Convert.ToDecimal(dataValue).ToString("N", new CultureInfo("en-US")));
                                }
                                else
                                {
                                    //Blank 0 data value will be added
                                    jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, 0);
                                }
                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[1, j].Value.ToString().Contains("E"))
                                {
                                    if (workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Formula != null)
                                    {
                                        jObject.Add(workSheet.Cells[1, j].Value.ToString() + j.ToString() + "_Formula", workSheet.Cells[i, j].Formula.ToString());
                                    }
                                    else if (workSheet.Cells[i, j].Formula != null)
                                    {
                                        jObject.Add(workSheet.Cells[1, j].Value.ToString() + j.ToString() + "_Formula", workSheet.Cells[i, j].Formula.ToString());
                                    }
                                }
                            }
                        }
                        else if (j == totalCols - 1)
                        {
                            if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                            {
                                if (workSheet.Cells[i, j].Value != null)
                                {
                                    var dataValue = workSheet.Cells[i, j].Value.ToString();//data values fetch
                                    decimal result = 0;
                                    bool canConvert = decimal.TryParse(dataValue, out result);
                                    if (canConvert == true)
                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, Convert.ToDecimal(dataValue).ToString("N", new CultureInfo("en-US")));
                                }
                                else
                                {
                                    //Blank 0 data value will be added
                                    jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, 0);
                                }
                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[1, j].Value.ToString().Contains("E"))
                                {
                                    if (workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Formula != null)
                                    {
                                        jObject.Add(workSheet.Cells[1, j].Value.ToString() + j.ToString() + "_Formula", workSheet.Cells[i, j].Formula.ToString());
                                    }
                                    else if (workSheet.Cells[i, j].Formula != null)
                                    {
                                        jObject.Add(workSheet.Cells[1, j].Value.ToString()+j.ToString() + "_Formula", workSheet.Cells[i, j].Formula.ToString());
                                    }
                                }
                            }
                        }
                    }

                }
                if (jObject.HasValues)
                    jArray.Add(jObject);
            }

            return jArray;
        }
        private static JArray GetExcelThemeArray(ExcelWorksheet workSheet, int totalRows)
        {
            try
            {
                JArray jthemeArray = new JArray();
                for (int i = 4; i <= totalRows; i++)
                {
                    JObject jObject = new JObject();
                    if (workSheet.Cells[i, 1].Value != null)
                    {
                        var LineItemText = workSheet.Cells[i, 1].Value.ToString();
                        if (LineItemText.ToLower().Trim().Equals("check"))
                        {
                            continue;
                        }

                        if (workSheet.Cells[i, 2].Value != null && workSheet.Cells[i, 2].Style.Font.Bold)
                            jObject.Add("config_bold", "bold");

                        if (workSheet.Cells[i, 2].Value != null && workSheet.Cells[i, 2].Style.Font.Italic)
                            jObject.Add("config_font", "italic");
                        if (workSheet.Cells[i, 2].Value != null && workSheet.Cells[i, 2].Style.Font.UnderLine)
                            jObject.Add("config_underline", "underline");
                        jObject.Add("LineItem", LineItemText);
                        if (workSheet.Cells[i, 1].Style.Font.Bold)
                            jObject.Add(workSheet.Cells[i, 1].Value.ToString(), "bold");

                        int totalCols = workSheet.Dimension.Columns;
                        for (int j = 2; j <= totalCols; j++)
                        {
                            if (j < totalCols - 1)
                            {
                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                {
                                    if (workSheet.Cells[i, j].Style.Font.Bold)
                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, "bold");
                                    if (workSheet.Cells[i, j].Style.Font.Italic)
                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_font", "italic");
                                    if (workSheet.Cells[i, j].Style.Font.UnderLine)
                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_underline", "underline");

                                    jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_color", workSheet.Cells[i, j].Style.Font.Color.Rgb);
                                    jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_size", workSheet.Cells[i, j].Style.Font.Size);
                                    if (workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Text.ToString().Contains('$'))
                                    {
                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_isdollar", "dollar");
                                    }
                                    if (workSheet.Cells[i, j].Value != null && workSheet.Cells[i, j].Value.ToString().Contains('%'))
                                    {
                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_ispercentage", "percentage");
                                    }

                                }
                            }
                            else if (j == totalCols - 1)
                            {
                                if (workSheet.Cells[1, j].Value != null && workSheet.Cells[2, j].Text != null)
                                {
                                    if (workSheet.Cells[i, j].Value != null)
                                    {
                                        if (workSheet.Cells[i, j].Style.Font.Bold)
                                            jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text, "bold");
                                        if (workSheet.Cells[i, j].Style.Font.Italic)
                                            jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_font", "italic");
                                        if (workSheet.Cells[i, j].Style.Font.UnderLine)
                                            jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_underline", "underline");

                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_color", workSheet.Cells[i, j].Style.Font.Color.Rgb);
                                        jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_size", workSheet.Cells[i, j].Style.Font.Size);
                                        if (workSheet.Cells[i, j].Text.ToString().Contains('$'))
                                        {
                                            jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_isdollar", "dollar");
                                        }
                                        if (workSheet.Cells[i, j].Text.ToString().Contains('%'))
                                        {
                                            jObject.Add(workSheet.Cells[1, j].Value + "<br/>" + workSheet.Cells[2, j].Text + "_ispercentage", "percentage");
                                        }

                                    }
                                }
                            }
                        }

                    }
                    if (jObject.HasValues)
                        jthemeArray.Add(jObject);
                }

                return jthemeArray;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }

}
