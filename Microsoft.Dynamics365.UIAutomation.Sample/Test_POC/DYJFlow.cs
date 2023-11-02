// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using TechTalk.SpecFlow.EnvironmentAccess;
using System.ComponentModel;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class DYJFlows
    {

        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestCategory("POC")]
        [TestMethod]
        public void DYJFlow()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
            {
                var instituteName = "Institute " + TestSettings.GetRandomString(7, 15) + " " + TestSettings.GetRandomString(3, 4) + "TA";
                var apprenticeName = "Apprentice " + TestSettings.GetRandomString(3, 10);
                var gurdianName = "Gurdian " + TestSettings.GetRandomString(3, 10);
                var contractName = "Contract " + TestSettings.GetRandomString(3, 10);

                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                xrmApp.Navigation.OpenApp(UCIAppName.DYJ);
                
                // Create Account
                xrmApp.Navigation.OpenSubArea("Group1", "Accounts");

                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.Entity.SetValue("name", instituteName);
                xrmApp.Entity.SetValue("websiteurl", "www." + TestSettings.GetRandomString(7, 15) + "nopagefountd.test.com");

                    //details
                xrmApp.Entity.SelectTab("Details");
                xrmApp.Entity.SetValue("description", "this is the text in the text area from automation" + TestSettings.GetRandomString(7, 15));

                xrmApp.CommandBar.ClickCommand("Save & Close");
                xrmApp.ThinkTime(3000);
                
                // Create Apprentice
                xrmApp.Navigation.OpenSubArea("Group1", "Apprentices");
                
                xrmApp.CommandBar.ClickCommand("New");
                xrmApp.Entity.SetValue("name", apprenticeName);
                
                xrmApp.Entity.SetValue("institute", instituteName);
                xrmApp.Lookup.OpenRecord(0);
                xrmApp.ThinkTime(2000);

                var birthDate = new DateTimeControl("dateofbirth") { Value = DateTime.Now.AddYears(-20) };
                xrmApp.Entity.SetValue(birthDate);

                var registerDate = new DateTimeControl("dateofregistration") { Value = DateTime.Now.AddDays(20) };
                xrmApp.Entity.SetValue(registerDate);

                xrmApp.Entity.SetValue(new OptionSet { Name = "gender", Value = "Male" });

                xrmApp.Entity.SetValue("address", "123 Brisbane City, QLD, 4000");
                xrmApp.CommandBar.ClickCommand("Save & Close");

                // Create Gurdian
                xrmApp.Navigation.OpenSubArea("Group1", "Gurdians");

                xrmApp.CommandBar.ClickCommand("New");
                xrmApp.Entity.SetValue("name", gurdianName);

                xrmApp.Entity.SetValue("apprenticename", apprenticeName);
                xrmApp.Lookup.OpenRecord(0);

                xrmApp.CommandBar.ClickCommand("Save & Close");
                
                // Create Contract Trainings
                xrmApp.Navigation.OpenSubArea("Group1", "ContractTrainings");

                xrmApp.CommandBar.ClickCommand("New");
                xrmApp.Entity.SetValue("contractname", contractName);

                var commenceDate = new DateTimeControl("commencedate") { Value = DateTime.Now };
                xrmApp.Entity.SetValue(commenceDate);

                xrmApp.Entity.SetValue("apprentice", apprenticeName);
                xrmApp.Lookup.OpenRecord(0);

                xrmApp.CommandBar.ClickCommand("Save & Close");





            }

        }
    }
}