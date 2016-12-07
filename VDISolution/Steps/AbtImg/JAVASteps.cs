using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtImg.Steps.AbtImg
{
    [Binding]
    public class JAVASteps
    {
        [Given(@"i opened the windows CMD and input java -version")]
        public void GivenIOpenedTheWindowsCMDAndInputJava_Version()
        {
            JAVA.proc.Start();
        }
        
        [Then(@"i can cofirm the Java Version is correct")]
        public void ThenICanCofirmTheJavaVersionIsCorrect()
        {
            Assert.True(JAVA.CheckVersion());
        }
    }
}
