using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtRegressionTest.Steps.AbtImg
{
    [Binding]
    public class AbtImgSteps
    {
        [Given(@"im using an Abt Computer Bit should be available and running at start up")]
        public void GivenImUsingAnAbtComputerBitShouldBeAvailableAndRunningAtStartUp()
        {
           Assert.True(Bit9.isParityRunning());
        }
    }
}
