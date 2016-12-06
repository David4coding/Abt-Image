using System;
using ImgDataModel;
using Xunit;

namespace AbtRegressiunTest
{
    public class IMG
    {   //[Fact]
        public void WinZipTest()
        {
            Assert.True(WinZip.addToZip());
        }
 //       [Fact]
        public void JAVATest()
        {
            Assert.True(JAVA.CheckVersion());
        }
      //  [Fact]
        public void SymantecTest()
        {
            Assert.True(Symantec.isSymantecActive());
        }
  //      [Fact]
        public void SilverLightTest()
        {
            Assert.True(SilverLight.isInstalled());
        }
    //    [Fact]
        public void FlashPlayerTest()
        {
            Assert.True(FlashPlayer.isInstalled() );
        }
    //    [Fact]
        public void LyncTest()
        {
            Lync.init();
            Assert.True(Lync.isLoggedIn());
        }
    //    [Fact]
        public void WordTest()
        {
         //  Assert.True( Office.WordWrapper.CreateWord());
        }
     //   [Fact]
        public void ExcelTest()
        {
       //     Assert.True(Office.ExcelWrapper.CreateExcel());
        }
      //  [Fact]
        public void PowerPointTest()
        {
            Assert.True(Office.PowerPointWrapper.CreatePowerPoint());
        }
      //  [Fact]
        public void AccessDataBaseTest()
        {
            Assert.True(Office.AccessWrapper.CreateAccess());
        }
       // [Fact]
        public void OutLookTest()
        {
            Assert.True(Office.OutlookWrapper.isLoggedIn());
        }
    }
}
