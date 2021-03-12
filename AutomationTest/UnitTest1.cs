using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections;
using System.Threading;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO;

namespace AutomationTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod2()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            var driver = new ChromeDriver("./", options);
            driver.Navigate().GoToUrl("https://www.udemy.com/course/objectivec/");

            Thread.Sleep(5000);

            var test = driver.FindElement(By.ClassName("curriculum--sub-header--23ncD"));
            test.FindElement(By.CssSelector("button")).Click();


            var Course = driver.FindElement(By.CssSelector("h1")).Text;
            IList<IWebElement> some = driver.FindElements(By.ClassName("section--panel--1tqxC"));
           
            //excel
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Test");

                var headerRow = new List<string[]>()
            {
                new string[] { "Course", "Section Name", "Lession Name", "Time" }
            };
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                var worksheet = excel.Workbook.Worksheets["Test"];
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                worksheet.Cells[headerRange].Style.Font.Bold = true;
                worksheet.Cells[headerRange].Style.Font.Size = 14;
                worksheet.Cells[headerRange].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                int recordIndex = 2;
                foreach (var element in some)
                {
                    var title = element.FindElement(By.ClassName("section--section-title--8blTh")).Text;



                    foreach (IWebElement cours in element.FindElements(By.ClassName("udlite-block-list-item-content"))) {
                        var time = cours.FindElements(By.TagName("span"));
                        string txtCurso = time[0].Text;
                        string txtTiempo = time[time.Count - 1].Text;
                        worksheet.Cells[recordIndex, 1].Value = Course;
                        worksheet.Cells[recordIndex, 2].Value = title;
                        worksheet.Cells[recordIndex, 3].Value = txtCurso;
                        worksheet.Cells[recordIndex, 4].Value = txtTiempo;
                        recordIndex++;
                    }
                }
                worksheet.Column(1).AutoFit();
                worksheet.Column(2).AutoFit();
                worksheet.Column(3).AutoFit();
                worksheet.Column(4).AutoFit();
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                FileInfo ef = new FileInfo(@""+ filePath + "\\" +Course+".xlsx");
                excel.SaveAs(ef);
                excel.Dispose();
            }
            //fin excel

            

            driver.Quit();

            Assert.AreEqual(1, 1);
        }
    }
}
