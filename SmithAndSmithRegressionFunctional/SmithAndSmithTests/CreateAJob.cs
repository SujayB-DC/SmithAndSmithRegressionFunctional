using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using System.Configuration;
using OpenQA.Selenium;

namespace Microsoft.Dynamics365.UIAutomation.Sample
{
    [TestClass]
    public class CreateAJob
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void CreateANewJob()
        {
            using (var xrmBrowser = new Microsoft.Dynamics365.UIAutomation.Api.Browser(BrowserType.Chrome))
            {
                xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);

                xrmBrowser.GuidedHelp.CloseGuidedHelp();

                xrmBrowser.ThinkTime(500);
                xrmBrowser.Navigation.OpenSubArea("Sales", "Jobs");

                xrmBrowser.ThinkTime(3000);
                xrmBrowser.CommandBar.ClickCommand("New");
                xrmBrowser.ThinkTime(2000);

                // test data source
                //var jobData = ExcelDataAccess.GetTestData("TestCreateNewJob");
                xrmBrowser.Entity.SetValue("dsl_suburbid", "0110 - Avenues");//jobData.Suburb);
                xrmBrowser.Driver.FindElement(By.Id("dsl_suburbid")).SendKeys(Keys.Tab);
                xrmBrowser.Entity.SetValue("dsl_servicelocationid_lookupValue", "Whangarei");//jobData.ServiceLocation);
                xrmBrowser.Driver.FindElement(By.Id("dsl_servicelocationid_lookupValue")).SendKeys(Keys.Tab);
                xrmBrowser.Entity.SetValue("dsl_accountid", "5611");//jobData.Account);
                xrmBrowser.Driver.FindElement(By.Id("dsl_accountid")).SendKeys(Keys.Tab);
                xrmBrowser.Entity.SetValue("dsl_registrationnumber", "HPZ495");// jobData.Rego);
                xrmBrowser.Driver.FindElement(By.Id("dsl_registrationnumber")).SendKeys(Keys.Tab);
                xrmBrowser.ThinkTime(2000);

                // check if another iframe exisits
                int size = xrmBrowser.Driver.FindElements(By.TagName("iframe")).Count;
                if (size > 1)
                {
                    // if another frame exisits, then need to close the dialog
                    xrmBrowser.Driver.SwitchTo().DefaultContent();
                    xrmBrowser.ThinkTime(2000);
                    if (xrmBrowser.Driver.IsVisible(By.Id("alertJs-divWarning")))
                    {
                        xrmBrowser.Driver.ClickWhenAvailable(By.Id("alertJs-close"));
                        // xrmBrowser.Driver.FindElement(By.XPath(Elements.Xpath[Reference.Login.StaySignedIn])).Submit();
                    }
                }
                xrmBrowser.ThinkTime(3000);
                xrmBrowser.CommandBar.ClickCommand("Save & Close");
                xrmBrowser.ThinkTime(2000);

            }
        }
    }
}
