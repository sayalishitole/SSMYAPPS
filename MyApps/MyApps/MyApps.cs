using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using OpenQA.Selenium;
using System.Drawing.Imaging;
using System.Configuration;

namespace MyApps
{
    [TestClass]
    public class MyApps

    {
        public static string siteUrl = ConfigurationManager.AppSettings["SiteUrl"];
        public static string username = ConfigurationManager.AppSettings["UserName"];
        public static string password = ConfigurationManager.AppSettings["Password"];
        //test 2
        ChromeDriver driver = new ChromeDriver("F:\\Chromedriver");

        [TestMethod]
        public void AddApps()
        {
            using (DataTable dtTemp = ExcelToDataTable("C:\\Users\\SayaliShitole\\Documents\\Visual Studio 2015\\Projects\\MyApps1.xlsx"))

            {
                
                
                 driver.Navigate().GoToUrl(siteUrl);

                    driver.FindElement(By.XPath("//*[@id='cred_userid_inputtext']")).SendKeys(username);
                    driver.FindElement(By.XPath("//*[@id='cred_password_inputtext']")).SendKeys(password);
                    Thread.Sleep(2000);
                    driver.FindElement(By.XPath("//*[@id='cred_sign_in_button']")).Click();
                    Thread.Sleep(TimeSpan.FromSeconds(5));
                    

                    

                   foreach (DataRow dRow1 in dtTemp.Rows)
                    {

                        var url = dRow1["Url"].ToString();
                        var Title = dRow1["Title"].ToString();
                        var AppUrl = dRow1["AppUrl"].ToString();
                        var AppIcon = dRow1["AppIcon"].ToString();
                        var newformUrl1 = dRow1["newformUrl"].ToString();
                      
                        var Addtitle = driver.FindElementByCssSelector("input[class$=TextField-field]");
                        //var Addtitle = driver.FindElement(By.XPath("body/script/styl/div data-bind/div id /div class / div class/ div class /div class /div class /div class /div class /nav class/div class /div class /div class /div class /label class /div data-bind/div class /div class /input[@value='od-TextEditor-input ms-TextField-field']"));
                        Addtitle.SendKeys(Title);

                        var AddappUrl = driver.FindElement(By.XPath("//input[@placeholder='Enter a URL']"));
                        // var AddappUrl = driver.FindElement(By.LinkText("Enter a URL"));
                        AddappUrl.SendKeys(AppUrl);

                        //var savebutton = driver.FindElement(By.XPath("//*[@id='appRoot']/div/div[2]/div[3]/div[2]/div[1]/div/div/div/div[2]/div[2]/button[1]/span"));
                        var savebutton = driver.FindElement(By.XPath("//*[@id='appRoot']/div/div[2]/div/div/div[3]/div[2]/div[1]/div/div/div/div[2]/div[2]/button[1]/span"));
                        savebutton.Click();
                        //Thread.Sleep(20000);
                        Thread.Sleep(TimeSpan.FromSeconds(5));
                        driver.Navigate().GoToUrl(newformUrl1);
                        Thread.Sleep(TimeSpan.FromSeconds(5));
                    }
                    //  }

                    ITakesScreenshot screenshotDriver = driver as ITakesScreenshot;
                    Screenshot screenshot = screenshotDriver.GetScreenshot();
                    String fp = "F:\\" + "snapshot" + "_" + DateTime.Now.ToString("dd_MMMM_hh_mm_ss_tt") + ".png";
                    screenshot.SaveAsFile(fp, ScreenshotImageFormat.Png);
                    // driver.Navigate().GoToUrl(newformUrl);

                    var Newbtn = driver.FindElement(By.XPath("//*[@id='appRoot']/div/div[2]/div/div/div[3]/div[1]/div/div[3]/div/div[2]/div[1]/div/div/span/i"));
                    Newbtn.Click();
                    var Appct = driver.FindElement(By.XPath("//*[@id='appRoot']/div/div[5]/div/div[1]/div/div/div[2]/div[1]/a/i"));
                    Appct.Click();
              
            }
        }
        public static DataTable ExcelToDataTable(string filePath)
        {
            DataTable dtexcel = new DataTable();
            bool hasHeaders = true;
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn;
            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
            else
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            //Looping Total Sheet of Xl File
            /*foreach (DataRow schemaRow in schemaTable.Rows)
            {
            }*/
            //Looping a first Sheet of Xl File
            DataRow schemaRow = schemaTable.Rows[0];
            string sheet = schemaRow["TABLE_NAME"].ToString();
            if (!sheet.EndsWith("_"))
            {
                string query = "SELECT  * FROM [" + sheet + "]";
                OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
                dtexcel.Locale = CultureInfo.CurrentCulture;
                daexcel.Fill(dtexcel);
            }

            conn.Close();
            return dtexcel;

        }
    }
}