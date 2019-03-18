using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.Reflection;
using System.Threading;

namespace AzureDevOpsTokenRevoker
{
    internal class Program
    {
        private const string userName = "INSERT_USER_NAME";
        private const string password = "INSERT_PASSWORD";
        private const string settingsUrl = "https://YOURORG.visualstudio.com/_usersSettings/tokens";

        private static void Main(string[] args)
        {
            var driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            driver.Manage().Timeouts().ImplicitWait = new TimeSpan(0, 0, 2);
            driver.Manage().Window.Minimize();
            driver.Navigate().GoToUrl(settingsUrl);

            DoLogin(driver);
            TrySettingSortOrder(driver);
            RevokeTokens(driver);
            driver.Close();
            driver.Dispose();
            Console.WriteLine("All done. Press any key to continue");
            Console.Read();
        }

        private static void DoLogin(ChromeDriver driver)
        {
            try
            {
                var inpMail = driver.FindElementById("i0116");
                inpMail.SendKeys(userName);
                var btnNext = driver.FindElementById("idSIButton9");
                btnNext.Click();
                Thread.Sleep(200); // Needed for Microsoft sign-in screen to do DOM changes
                var inpPassword = driver.FindElementById("i0118");
                inpPassword.SendKeys(password);
                btnNext = driver.FindElementById("idSIButton9");
                btnNext.Click();
                Thread.Sleep(200); // Needed for Microsoft sign-in screen to do DOM changes
                HandleSMSPasscode(driver);
                btnNext = driver.FindElementById("idBtn_Back");
                btnNext.Click();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static void HandleSMSPasscode(ChromeDriver driver)
        {
            try
            {
                //Check if SMS code is neccesary
                var inpSmscode = driver.FindElementById("idTxtBx_SAOTCC_OTC");
                Console.WriteLine("Please enter SMS passcode:");
                var smscode = Console.ReadLine();
                Console.WriteLine("Continuing login process...");
                inpSmscode.SendKeys(smscode);
                var btnLogin = driver.FindElementById("idSubmit_SAOTCC_Continue");
                btnLogin.Click();
            }
            catch (NoSuchElementException)
            {
                //Expected when multi-factor authentication is not activated...
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private static void TrySettingSortOrder(ChromeDriver driver)
        {
            try
            {
                var btnSortBy = driver.FindElementByCssSelector("span[id$='-expiresOn']");
                btnSortBy.Click();
                btnSortBy = driver.FindElementByCssSelector("span[id$='-expiresOn']");
                btnSortBy.Click();
            }
            catch (Exception)
            {
                //Continue without sorting...
            }
        }

        private static void RevokeTokens(ChromeDriver driver)
        {
            var token = GetTokenElement(driver);
            while (token != null)
            {
                try
                {
                    token.Click();
                    var btnRewoke = driver.FindElementByXPath("/html/body/div[1]/div/div/div[3]/div/div[2]/div/div[2]/div[1]/button[2]");
                    btnRewoke.Click();
                    var btnConfirm = driver.FindElementByXPath("/html/body/div[2]/div/div/div[2]/div/div[3]/div/button[1]");
                    btnConfirm.Click();
                    Thread.Sleep(500);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                token = GetTokenElement(driver);
            }
        }

        private static IWebElement GetTokenElement(ChromeDriver driver)
        {
            IWebElement token;
            try
            {
                token = driver.FindElementByXPath("/html/body/div[1]/div/div/div[3]/div/div[2]/div/div[3]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/div[1]");
            }
            catch (NoSuchElementException)
            {
                //Expected to occur when done so no printing to console...
                token = null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                token = null;
            }
            return token;
        }
    }
}
