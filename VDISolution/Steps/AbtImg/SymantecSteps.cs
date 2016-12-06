using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace VDISolution.Steps.AbtImg
{
    [Binding]
    public class SymantecSteps
    {
        [Given(@"Symantec active scan is running")]
        public void GivenSymantecActiveScanIsRunning()
        {
            Assert.True(Symantec.isSymantecActive());
        }
    }
}
