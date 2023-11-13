using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System;
using TestContext = Microsoft.VisualStudio.TestTools.UnitTesting.TestContext;
using OpenQA.Selenium.DevTools.V113.Tethering;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics.Eventing.Reader;
using AventStack.ExtentReports.MarkupUtils;
//using AventStack.ExtentReports.Reporter.Config;


namespace Microsoft.Dynamics365.UIAutomation.Sample.Test_POC

{
    [TestClass]
    public class Tests
    {
        protected static ExtentReports Extent;
        protected static ExtentTest TestParent;
        protected static ExtentTest Test;
        protected static string AssemblyName;
        public TestContext TestContext { get; set; }

        [AssemblyInitialize]
        public static void AssemblyInitialize(TestContext context)
        {
            AssemblyName = Assembly.GetExecutingAssembly().GetName().Name;

            // http://extentreports.com/docs/versions/4/net/
            var dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "\\";
            const string fileName = "ExtentReport.html";
            var htmlReporter = new ExtentHtmlReporter(dir + fileName);
          
            /**  htmlReporter.Config.DocumentTitle = $"Test Results: {DateTime.Now:MM/dd/yyyy h:mm tt}";
            htmlReporter.Config.ReportName = context.FullyQualifiedTestClassName;
            htmlReporter.Config.Theme = Theme.Dark;
          **/
            // Add any additional contextual information
            Extent = new ExtentReports();
            Extent.AddSystemInfo("Browser", Enum.GetName(typeof(BrowserType), TestSettings.Options.BrowserType));
            Extent.AddSystemInfo("Test User",
                System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"]);
            Extent.AddSystemInfo("D365 CE Instance",
                System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"]);
            Extent.AttachReporter(htmlReporter);
           context.AddResultFile(fileName);

            // Create a container for the tests in the class
            TestParent = Extent.CreateTest(context.FullyQualifiedTestClassName);
        }
        [TestInitialize]
        public void TestInitialize()
        {
            // Get unit test description attribute
            var type = Type.GetType($"{TestContext.FullyQualifiedTestClassName}, {AssemblyName}");
            var methodInfo = type?.GetMethod(TestContext.TestName);
            var customAttributes = methodInfo?.GetCustomAttributes(false);
            DescriptionAttribute desc = null;
            if (customAttributes != null)
            {
                foreach (var n in customAttributes)
                {
                    desc = n as DescriptionAttribute;
                    if (desc != null)
                        break;
                }
            }

            // Create individual test under the parent container / class
            Test = TestParent.CreateNode(TestContext.TestName, desc?.Description);
        }


        [TestCleanup]
        public void TestCleanup()
        {

            // Sets individual Extent test result so it reflects correctly in the report
           
            if (Test.Status == Status.Error)
                return;

            switch (TestContext.CurrentTestOutcome)
            {
                case UnitTestOutcome.Error:
                    Test.Fail("Test Failed - System Error");
                    break;
                case UnitTestOutcome.Passed:
                    Test.Pass("Test Passed");
                    break;
                case UnitTestOutcome.Failed:
                    Test.Fail("Test Failed");
                    break;
                case UnitTestOutcome.Inconclusive:
                    Test.Fail("Test Failed - Inconclusive");
                    break;
                case UnitTestOutcome.Timeout:
                    Test.Fail("Test Failed - Timeout");
                    break;
                //case UnitTestOutcome.NotRunnable:
                case UnitTestOutcome.Aborted:
                    Test.Skip("Test Failed - Aborted / Not Runnable");
                    break;
                case UnitTestOutcome.InProgress:
                case UnitTestOutcome.Unknown:
                default:
                    Test.Fail("Test Failed - Unknown");
                    break;
            }
            
        }
        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            Extent.Flush();
        }
        public void LogExceptionAndFail(Exception e)
        {
            // Formats the exception details to look nice
            var message = e.Message + Environment.NewLine + e.StackTrace.Trim();
            var markup = MarkupHelper.CreateCodeBlock(message);
            Test.Error(markup);
            throw e;
        }
    }
}