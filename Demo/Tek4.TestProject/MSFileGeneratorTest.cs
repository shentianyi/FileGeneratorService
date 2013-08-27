using Tek4.Highcharts.Exporting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Svg;
using System.Collections.Generic;
using System.IO;

namespace Tek4.TestProject
{
    
    
    /// <summary>
    ///这是 MSFileGeneratorTest 的测试类，旨在
    ///包含所有 MSFileGeneratorTest 单元测试
    ///</summary>
    [TestClass()]
    public class MSFileGeneratorTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///获取或设置测试上下文，上下文提供
        ///有关当前测试运行及其功能的信息。
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

        #region 附加测试特性
        // 
        //编写测试时，还可使用以下特性:
        //
        //使用 ClassInitialize 在运行类中的第一个测试前先运行代码
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //使用 ClassCleanup 在运行完类中的所有测试后再运行代码
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //使用 TestInitialize 在运行每个测试前先运行代码
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //使用 TestCleanup 在运行完每个测试后运行代码
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///CreateDoc 的测试
        ///</summary>
        [TestMethod()]
        public void CreateDocTest()
        {
            MSFileGenerator target = new MSFileGenerator(); // TODO: 初始化为适当的值
          
            Assert.Inconclusive("无法验证不返回值的方法。");
        }

        /// <summary>
        ///CreateDocStream 的测试
        ///</summary>
        [TestMethod()]
        public void CreateDocStreamTest()
        {  
        }

        /// <summary>
        ///CreateExcelXStream 的测试
        ///</summary>
        [TestMethod()]
        public void CreateExcelXStreamTest()
        {
            MSFileGenerator target = new MSFileGenerator(); // TODO: 初始化为适当的值
            List<SvgDocument> svgDocs = null; // TODO: 初始化为适当的值
            Stream stream = null; // TODO: 初始化为适当的值
            target.CreateExcelXStream(svgDocs, stream);
            Assert.Inconclusive("无法验证不返回值的方法。");
        }

        /// <summary>
        ///CreateExcelStream 的测试
        ///</summary>
        [TestMethod()]
        public void CreateExcelStreamTest()
        {
            MSFileGenerator target = new MSFileGenerator(); // TODO: 初始化为适当的值
            List<SvgDocument> svgDocs = null; // TODO: 初始化为适当的值
            Stream stream = null; // TODO: 初始化为适当的值
            target.CreateExcelStream(svgDocs, stream);
            Assert.Inconclusive("无法验证不返回值的方法。");
        }
    }
}
