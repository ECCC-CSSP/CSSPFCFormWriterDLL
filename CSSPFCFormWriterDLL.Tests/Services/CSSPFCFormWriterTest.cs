using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using CSSPFCFormWriterDLL.Services;
using CSSPLabSheetParserDLL.Services;
using System.Globalization;
using CSSPModelsDLL.Models;
using CSSPEnumsDLL.Enums;

namespace CSSPFCFormWriterDLL.Tests.Services
{
    /// <summary>
    /// Summary description for CSSPFCFormWriterTest
    /// </summary>
    [TestClass]
    public class CSSPFCFormWriterTest
    {
        #region Variables
        private TestContext testContextInstance;
        #endregion Variables

        #region Properties
        public CSSPLabSheetParser csspLabSheetParser { get; set; }
        public CSSPFCFormWriter csspFCFormWriter { get; set; }
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
        #endregion Properties

        #region Constructors
        public CSSPFCFormWriterTest()
        {
        }
        #endregion Constructors

        #region Initialize and Cleanup
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
        #endregion Initialize and Cleanup

        #region Testing functions
        [TestMethod]
        public void LabSheetService_CreateFCForm_Good_Test()
        {
            foreach (LanguageEnum LanguageRequest in new List<LanguageEnum>() { LanguageEnum.en }) //, "fr" })
            {
                Setup(LanguageRequest);

                string retStr = csspFCFormWriter.CreateFCForm(@"C:\Users\leblancc\Desktop\FCFormTest.docx");
                Assert.AreEqual("", retStr);
            }
        }
        #endregion Testing functions

        #region Functions private
        public void Setup(LanguageEnum LanguageRequest)
        {
            CultureInfo cultureInfo = new CultureInfo(LanguageRequest + "-CA");
            csspLabSheetParser = new CSSPLabSheetParser();

            FileInfo fiLabSheetTestFile = new FileInfo(@"C:\CSSP latest code\CSSPLabSheetParserDLL\CSSPLabSheetParserDLLTest\LabSheetTestFile.txt");
            Assert.IsTrue(fiLabSheetTestFile.Exists);

            StreamReader sr = fiLabSheetTestFile.OpenText();
            string FullFileText = sr.ReadToEnd();
            sr.Close();

            LabSheetA1Sheet labSheetA1Sheet = csspLabSheetParser.ParseLabSheetA1(FullFileText);
            Assert.AreEqual("", labSheetA1Sheet.Error);

            csspFCFormWriter = new CSSPFCFormWriter(LanguageRequest, FullFileText);
        }
        #endregion Functions private
    }
}
