using System;
using Xunit;

namespace GoogleDriveTest
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            var files = Satrex.GoogleDrive.GoogleDriveInternal.ListFiles("1r-wagtc1rNhkkIbu8Mcw06WYWAM1kdQf");

            Assert.NotEmpty(files);

        }
    }
}
