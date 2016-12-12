using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtRegressionTest.Steps.AbtImg
{
    [Binding]
    public class CarbonBlackSteps
    {
        [Given(@"im using an Abt Computer Carbon Black should be available and running at start up")]
        public void GivenImUsingAnAbtComputerCarbonBlackShouldBeAvailableAndRunningAtStartUp()
        {
            Assert.True(CarbonBlack.isCarbonBlackProcessRunning());
        }
    }
}
