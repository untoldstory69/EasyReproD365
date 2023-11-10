// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using TechTalk.SpecFlow.EnvironmentAccess;
using System.ComponentModel;

namespace Microsoft.Dynamics365.UIAutomation.Sample.Test_POC
{
    [TestClass]
    public class AccountContact
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestCategory("POC")]
        [TestMethod]
        public void AccContact()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                xrmApp.Navigation.OpenApp(UCIAppName.Sales);
                
                // Create Account
                xrmApp.Navigation.OpenSubArea("Sales", "Accounts");

                xrmApp.CommandBar.ClickCommand("New");

                var accountName = "Test Account" + TestSettings.GetRandomString(7,15) + " " + TestSettings.GetRandomString(3,4) + "TA";
                xrmApp.Entity.SetValue("name", accountName);
                xrmApp.Entity.SetValue("websiteurl", "www." + TestSettings.GetRandomString(7, 15) + "nopagefountd.test.com");

                //details
                xrmApp.Entity.SelectTab("Details");
                xrmApp.Entity.SetValue("description", "this is the text in the text area from automation" + TestSettings.GetRandomString(7, 15));

                xrmApp.CommandBar.ClickCommand("Save & Close");
                xrmApp.ThinkTime(3000);

                //Create Contact
                xrmApp.Navigation.OpenSubArea("Sales", "Contacts");

                xrmApp.CommandBar.ClickCommand("New");

                //summary
                xrmApp.Entity.SetValue("firstname", "Test FName" + TestSettings.GetRandomString(2,5) + " " + TestSettings.GetRandomString(3,4));
                xrmApp.Entity.SetValue("lastname", "Test LName" + TestSettings.GetRandomString(2,5)  + " " + TestSettings.GetRandomString(3, 4));
                xrmApp.Entity.SetValue("parentcustomerid", accountName);
                xrmApp.Lookup.OpenRecord(0);

                //details
                xrmApp.Entity.SelectTab("Details");
                xrmApp.Entity.SetValue(new OptionSet { Name = "gendercode", Value = "1" });
                
                var birthDate = new DateTimeControl("birthdate") { Value = DateTime.Now};
                xrmApp.Entity.SetValue(birthDate);
                xrmApp.Entity.Save();
                xrmApp.ThinkTime(3000);


            }

        }
    }
}