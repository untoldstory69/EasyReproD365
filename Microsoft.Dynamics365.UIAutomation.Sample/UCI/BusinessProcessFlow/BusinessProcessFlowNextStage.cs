﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using System;
using System.Security;
using Microsoft.Dynamics365.UIAutomation.Sample.Test_POC;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI
{
    [TestClass]
    public class BusinessProcessFlowNextStageUCI : ExtentReport
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        [TestCategory("BusinessProcessFlow")]
        public void UCITestBusinessProcessFlowNextStage()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))
                try
                {
                    {
                        xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                        xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                        xrmApp.Navigation.OpenSubArea("Sales", "Leads");

                        xrmApp.Grid.SwitchView("Open Leads");

                        xrmApp.Grid.OpenRecord(0);

                        xrmApp.BusinessProcessFlow.NextStage("Qualify");
                    }

                }

                 
                catch (Exception ex)
                {

                    LogExceptionAndFail(ex);
                    AddScreenShot(client, "Failed Screen");
                 }

}
    }
}
