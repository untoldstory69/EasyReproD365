// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Api;
using System.Diagnostics;
using System.Linq;

namespace Microsoft.Dynamics365.UIAutomation.Sample.Test_POC
{
    [TestClass]
    public class AccountContactCopy : Tests
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestCategory("POC")]
        [TestMethod]
        public void AccContactCopy()
        {
            try
            {
                var client = new WebClient(TestSettings.Options);
                using (var xrmApp = new XrmApp(client))
                {
                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    // Create Account
                    xrmApp.Navigation.OpenSubArea("Sales", "Accounts");
                    xrmApp.CommandBar.ClickCommand("New");

                    var accountName = "Test Account" + TestSettings.GetRandomString(7, 15) + " " + TestSettings.GetRandomString(3, 4) + "TA";
                    xrmApp.Entity.SetValue("name", accountName);
                    var name = xrmApp.Entity.GetValue("name");
                    //Assert.IsNotNull(name);
                    Assert.AreEqual("Failing test to see failed test result", name, "The text are not equal");

                }
            }
            catch (Exception ex) {
                LogExceptionAndFail(ex);
            }
            

        }
    }
}