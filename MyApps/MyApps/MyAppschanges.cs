using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using OpenQA.Selenium;

namespace MyApps
{
    [TestClass]
    public class MyApps1

    {


        [TestMethod]
        public void AddnewApps()
        {
            using (DataTable dtTemp = ExcelToDataTable("C:\\Users\\SayaliShitole\\Documents\\Visual Studio 2015\\Projects\\MyApps1.xlsx"))

            {
                foreach (DataRow dRow in dtTemp.Rows)
                {

                    var url = dRow["Url"].ToString();
                    var Username = dRow["Username"].ToString();
                    var Password = dRow["Password"].ToString();
                    var Title = dRow["Title"].ToString();
                    var AppUrl = dRow["AppUrl"].ToString();
                    var AppIcon = dRow["AppIcon"].ToString();
                    var newformUrl = dRow["newformUrl"].ToString();
                    ChromeDriver driver = new ChromeDriver("F:\\Chromedriver");

                    //var popup = driver.WindowHandles[1];

                    driver.Navigate().GoToUrl(newformUrl);
                    var UserName = driver.FindElement(By.XPath("//*[@id='cred_userid_inputtext']"));
                    //Below code will enter the password for the given user from excel.
                    var PassWord = driver.FindElement(By.XPath("//*[@id='cred_password_inputtext']"));
                    UserName.SendKeys(Username);
                    PassWord.SendKeys(Password);
                    Thread.Sleep(6000);
                    //Click on signin button
                    //  var MFA = driver.FindElement(By.XPath("//*[@id='aad_account_tile']"));
                    //MFA.Click();

                    var LoginButton = driver.FindElement(By.XPath("//*[@id='cred_sign_in_button']"));
                    LoginButton.Click();
                    Thread.Sleep(20000);
                   // var Managemyapps = driver.FindElement(By.XPath("//button[@class='btn btn-default']"));
                    //Thread.Sleep(6000);
                    //Managemyapps.Click();
                    //Thread.Sleep(6000);
                    //var Administrator = driver.FindElementByCssSelector(".btn btn-info ng-binding");
                    //var Administrator = driver.FindElement(By.LinkText("My Apps Administration"));
                    // var Savebutton = driver.FindElementsByCssSelector("input[id$=Default]")[0];
                    // driver.SwitchTo().Frame(0);

                    //Thread.Sleep(6000);
                    //Administrator.Click();
                    //Thread.Sleep(6000);


                  //Thread.Sleep(3000);
                    //    Thread.Sleep(6000);
                    //Assert.IsTrue(!string.IsNullOrEmpty(popup));
                    //Assert.AreEqual(driver.SwitchTo().Window(popup).Url, "https://instantintranet.sharepoint.com/sites/start/Lists/MyAppsGenNL/NewForm.aspx?Source=https%3A%2F%2Finstantintranet%2Esharepoint%2Ecom%2Fsites%2Fstart%2FLists%2FMyAppsGenNL%2FAllItems%2Easpx&ContentTypeId=0x0100014263E85C04C84D91E5796D847E3FB40030C1ACF46A0CDF4AA772E2D39395DBE7&RootFolder=");
                    //var NewItem = driver.FindElement(By.LinkText("CommandBarItem-link"));
                    //NewItem.Click();
                    // var AppCt = driver.FindElement(By.XPath("//*[@id='appRoot']/div/div[8]/div[2]/div/div/div[2]/div[1]/a"));



                    driver.Navigate().GoToUrl(newformUrl);

                    Thread.Sleep(6000);




                    //for (int i = 0; i <= 3; i++)
                    //{
                    var Addtitle = driver.FindElementByCssSelector("input[class$=TextField-field]");
                    //var Addtitle = driver.FindElement(By.XPath("body/script/styl/div data-bind/div id /div class / div class/ div class /div class /div class /div class /div class /nav class/div class /div class /div class /div class /label class /div data-bind/div class /div class /input[@value='od-TextEditor-input ms-TextField-field']"));
                    Addtitle.SendKeys(Title);

                    var AddappUrl = driver.FindElement(By.XPath("//input[@placeholder='Enter a URL']"));
                    // var AddappUrl = driver.FindElement(By.LinkText("Enter a URL"));
                    AddappUrl.SendKeys(AppUrl);

                    var savebutton = driver.FindElement(By.XPath("//*[@id='appRoot']/div/div[2]/div[3]/div[2]/div[1]/div/div/div/div[2]/button[1]"));
                    savebutton.Click();

                    Thread.Sleep(20000);
                   driver.Navigate().GoToUrl(newformUrl);



                    //var newitem = driver.FindElement(By.LinkText("New"));
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