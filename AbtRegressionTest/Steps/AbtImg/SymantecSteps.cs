using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtRegressionTest.Steps.AbtImg
{
    [Binding]
    public class SymantecSteps
    {
        [Given(@"Symantec active scan is running")]
        public void GivenSymantecActiveScanIsRunning()
        {
            Assert.True(Symantec.isSymantecActive());
        }

        [Given(@"im using an Abt Computer Encryption Desktop should be available and running at start up")]
        public void GivenImUsingAnAbtComputerEncryptionDesktopShouldBeAvailableAndRunningAtStartUp()
        {
            Assert.True(Symantec.findEncriptionRegistry());
        }

        [Given(@"im using an Abt Computer Encryption Desktop process running at start up")]
        public void GivenImUsingAnAbtComputerEncryptionDesktopProcessRunningAtStartUp()
        {
            Assert.True(Symantec.isEncriptionProcessRunning());
        }


    }
}
