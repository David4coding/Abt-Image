using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;


namespace AbtRegressionTest.Steps.AbtImg
{
    [Binding]
    public class SCCMSteps
    {
        [Given(@"SCCM is available at start up")]
        public void GivenSCCMIsAvailableAtStartUp()
        {
            Assert.True(SCCM.isAgentAvailable());
        }

    }
}

