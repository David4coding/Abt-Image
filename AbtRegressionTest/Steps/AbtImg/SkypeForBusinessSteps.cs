using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtImg.Steps.AbtImg
{
    [Binding]
    public class SkypeForBusinessSteps
    {
        [Given(@"i have logged to Windows")]
        public void GivenIHaveLoggedToWindows()
        {
            Lync.init();
        }
        
        [Then(@"Skype for Business opened with my user")]
        public void ThenSkypeForBusinessOpenedWithMyUser()
        { 
            Assert.True(Lync.isLoggedIn());
        }
    }
}
