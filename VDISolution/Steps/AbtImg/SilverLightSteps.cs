using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace VDISolution.Steps.AbtImg
{
    [Binding]
    public class SilverLightSteps
    {
        [Given(@"i see SilverLight is installed")]
        public void GivenISeeSilverLightIsInstalled()
        {
            SilverLight.findRegistry();
        }
        
        [Then(@"i can see Silverlight's registry")]
        public void ThenICanSeeSilverlightSRegistry()
        {
            Assert.True(SilverLight.checkSilverLightVersion());
        }
    }
}
