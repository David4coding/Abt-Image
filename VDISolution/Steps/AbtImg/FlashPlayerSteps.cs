using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace VDISolution.Steps.AbtImg
{
    [Binding]
    public class FlashPlayerSteps
    {
        [Given(@"i see FlashPlayer is installed")]
        public void GivenISeeFlashPlayerIsInstalled()
        {
            FlashPlayer.findRegistryKey();
        }
        
        [Then(@"i can see Flash's registry")]
        public void ThenICanSeeFlashSRegistry()
        {
            Assert.True(FlashPlayer.assertFlashPlayer());
        }
    }
}
