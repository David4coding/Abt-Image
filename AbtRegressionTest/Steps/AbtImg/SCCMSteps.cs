using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;


namespace AbtImg.Steps.AbtImg
{
    [Binding]
    public class AbtImgSteps
    {
        [Given(@"im using an abt laptop SCCM should be available and running at start up")]
        public void GivenImUsingAnAbtLaptopSCCMShouldBeAvailableAndRunningAtStartUp()
        {
            Assert.True(SCCM.isAgentAvailable());
        }
    }
}

