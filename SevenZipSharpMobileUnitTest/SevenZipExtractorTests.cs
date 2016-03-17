using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SevenZip;

namespace SevenZipSharpMobileUnitTest
{
    /// <summary>
    /// Summary description for SevenZipExtractorTests
    /// </summary>
    /// <remarks>
    ///          See A starter on smart device unit testing
    ///          http://blogs.msdn.com/b/prativen/archive/2007/12/13/a-starter-on-smart-device-unit-testing.aspx
    ///          
    ///          Overview of Smart Device Unit Tests
    ///          https://msdn.microsoft.com/en-us/library/bb513825(v=vs.90).aspx
    ///          
    ///          Smart Device Test Projects
    ///          https://msdn.microsoft.com/en-us/library/bb545995(v=vs.90).aspx
    ///          
    ///          How to: Debug while Running a Smart Device Unit Test:
    ///          https://msdn.microsoft.com/en-us/library/bb513875%28v=vs.90%29.aspx?f=255&MSPPError=-2147217396
    /// 
    /// </remarks>
    [TestClass]
    public class SevenZipExtractorTests
    {
        public SevenZipExtractorTests()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void CanUnzipSevenZipFileContainingOneFileNoDirectories()
        {
            var executingPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            var zipFileName = executingPath + "\\gpl.7z";
            var zipContentsFileName = executingPath + "\\gpl.txt";

            Assert.IsTrue(File.Exists(zipFileName), string.Format("Test setup failed, {0} file not found!", zipFileName));

            if(File.Exists(zipContentsFileName))
                File.Delete(zipContentsFileName);

            try
            {
                using (var extr = new SevenZipExtractor(zipFileName))
                {
                    try
                    {
                        Assert.IsFalse(File.Exists(zipContentsFileName));
                        extr.ExtractArchive(executingPath);
                        Assert.IsTrue(File.Exists(zipContentsFileName), string.Format("Unzipped file '{0}' not found!", zipContentsFileName));
                    }
                    catch (Exception exception)
                    {
                        throw;
                    }
                }
            }
            catch (Exception exception)
            {
                throw;
            }
        }
    }
}
