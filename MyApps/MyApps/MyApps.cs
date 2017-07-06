using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using OpenQA.Selenium;
using System.Drawing.Imaging;

namespace MyApps
{
    [TestClass]
    public class MyApps

    {
        //test 2
        ChromeDriver driver = new ChromeDriver("F:\\Chromedriver");

        [TestMethod]
        public void AddApps()
        {
            using (DataTable dtTemp = ExcelToDataTable("C:\\Users\\SayaliShitole\\Documents\\Visual Studio 2015\\Projects\\MyApps1.xlsx"))

            {
                foreach (DataRow dRow in dtTemp.Rows)
                {

                    //var url = dRow["Url"].ToString();
                    var Username = dRow["Username"].ToString();
                    var Password = dRow["Password"].ToString();
                    //var Title = dRow["Title"].ToString();
                    //var AppUrl = dRow["AppUrl"].ToString();
                    //var AppIcon = dRow["AppIcon"].ToString();
                    var newformUrl = dRow["newformUrl"].ToString();


                    //var popup = driver.WindowHandles[1];

                    driver.Navigate().GoToUrl(newformUrl);


                    //var ss = driver.GetScreenshot();
                    //ss.SaveAsFile("f:\\google"as , System.Drawing.Imaging.ImageFormat.Png);

                    var UserName = driver.FindElement(By.XPath("//*[@id='cred_userid_inputtext']"));
                    //Below code will enter the password for the given user from excel.
                    var PassWord = driver.FindElement(By.XPath("//*[@id='cred_password_inputtext']"));
                    UserName.SendKeys(Username);
                    PassWord.SendKeys(Password);
                    Thread.Sleep(2000);
                    //Click on signin button
                    //  var MFA = driver.FindElement(By.XPath("//*[@id='aad_account_tile']"));
                    //MFA.Click();

                    var LoginButton = driver.FindElement(By.XPath("//*[@id='cred_sign_in_button']"));
                    LoginButton.Click();
                    // Thread.Sleep(20000);
                    Thread.Sleep(TimeSpan.FromSeconds(5));
                    //  var Managemyapps = driver.FindElement(By.XPath("//button[@class='btn btn-default']"));
                    //Thread.Sleep(6000);
                    //Managemyapps.Click();
                    //Thread.Sleep(6000);
                    //var Administrator = driver.FindElementByCssSelector(".btn btn-info ng-binding");
                    // var Administrator = driver.FindElement(By.LinkText("My Apps Administration"));
                    // var Savebutton = driver.FindElementsByCssSelector("input[id$=Default]")[0];
                    // driver.SwitchTo().Frame(0);

                    //Thread.Sleep(6000);
                    //Administrator.Click();
                    //Thread.Sleep(3000);
                    //    Thread.Sleep(6000);
                    //Assert.IsTrue(!string.IsNullOrEmpty(popup));
                    //Assert.AreEqual(driver.SwitchTo().Window(popup).Url, "https://instantintranet.sharepoint.com/sites/start/Lists/MyAppsGenNL/NewForm.aspx?Source=https%3A%2F%2Finstantintranet%2Esharepoint%2Ecom%2Fsites%2Fstart%2FLists%2FMyAppsGenNL%2FAllItems%2Easpx&ContentTypeId=0x0100014263E85C04C84D91E5796D847E3FB40030C1ACF46A0CDF4AA772E2D39395DBE7&RootFolder=");
                    //var NewItem = driver.FindElement(By.LinkText("CommandBarItem-link"));
                    //NewItem.Click();
                    // var AppCt = driver.FindElement(By.XPath("//*[@id='appRoot']/div/div[8]/div[2]/div/div/div[2]/div[1]/a"));



                    // driver.Navigate().GoToUrl(newformUrl);

                    //Thread.Sleep(2000);




                    //for (int i = 0; i <= dtTemp.Rows.Count; i++)
                    //{
                    foreach (DataRow dRow1 in dtTemp.Rows)
                    {

                        var url = dRow1["Url"].ToString();
                        var Title = dRow1["Title"].ToString();
                        var AppUrl = dRow1["AppUrl"].ToString();
                        var AppIcon = dRow1["AppIcon"].ToString();
                        var newformUrl1 = dRow["newformUrl"].ToString();
                      
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


                    // var newtab1 =driver.FindElement("By.CssSelector('body')).SendKeys(Keys.CONTROL + "\t"");

                ////Commented by Shrikant
                    //var Addtitle1 = driver.FindElementByCssSelector("input[class$=TextField-field]");
                    ////var Addtitle = driver.FindElement(By.XPath("body/script/styl/div data-bind/div id /div class / div class/ div class /div class /div class /div class /div class /nav class/div class /div class /div class /div class /label class /div data-bind/div class /div class /input[@value='od-TextEditor-input ms-TextField-field']"));
                    //Addtitle1.SendKeys(Title);

                    //var AddappUrl1 = driver.FindElement(By.XPath("//input[@placeholder='Enter a URL']"));
                    //// var AddappUrl = driver.FindElement(By.LinkText("Enter a URL"));
                    //AddappUrl1.SendKeys(AppUrl);

                    //var savebutton1 = driver.FindElement(By.XPath("//*[@id='appRoot']/div/div[2]/div[3]/div[2]/div[1]/div/div/div/div[2]/button[1]"));
                    //savebutton1.Click();



                    //var newitem = driver.FindElement(By.XPath("//span[@class='od-IconGlyph ms-Icon ms-Icon--Add od-IconGlyph--visible']/class[2]"));
                    //newitem.Click();
                    //var Addappicon = driver.FindElement(By.XPath("//div[@class='od-TextEditor-input ms-TextField-field']"));
                    //Addtitle.SendKeys(AppIcon);
                    //}
                }
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