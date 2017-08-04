using Microsoft.VisualStudio.TestTools.UnitTesting;
using OutsuranceFilePrinter;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace OutsuranceFilePrinter.Tests
{
    [TestClass()]
    public class OutsuranceFilePrinterTests
    {
        [TestMethod()]
        public void SortAndPrintNameDataTest()
        {
            DataTable TestData = CreateTestData();

            OutsuranceFilePrinter opf = new OutsuranceFilePrinter();

            string outputPath = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.TestDir))
              + "\\OutsuranceFilePrinterTests\\bin\\Debug";

            opf.SortAndPrintNameData(TestData);

            List<string> Rows = new List<string>();

            using (StreamReader file = new StreamReader(outputPath + "\\Names.txt"))
            {


                while (file.Peek() >= 0)
                    Rows.Add(file.ReadLine());
            }

            Assert.AreEqual(9, Rows.Count);
            Assert.AreEqual(Rows[3], "Howe,2");
        }

        public TestContext TestContext { get; set; }

        [TestMethod()]
        public void SortAndPrintAddressDataTest()
        {
            DataTable TestData = CreateTestData();

            OutsuranceFilePrinter opf = new OutsuranceFilePrinter();


            string outputPath = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.TestDir))
                + "\\OutsuranceFilePrinterTests\\bin\\Debug";
           

            opf.SortAndPrintAddressData(TestData);

            List<string> Rows = new List<string>();

            using (StreamReader file = new StreamReader(outputPath + "\\Addresses.txt"))
            {
                while (file.Peek() >= 0)
                    Rows.Add(file.ReadLine());
            }

            Assert.AreEqual(8, Rows.Count);
            Assert.AreEqual(Rows[3], "102 Long Lane");
        }

        private DataTable CreateTestData()
        {
            DataTable TestData = new DataTable();
            TestData.Columns.Add(new DataColumn("FirstName"));
            TestData.Columns.Add(new DataColumn("LastName"));
            TestData.Columns.Add(new DataColumn("Address"));

            TestData.Rows.Add("Jimmy", "Smith", "102 Long Lane");
            TestData.Rows.Add("Clive", "Owen", "65 Ambling Way");
            TestData.Rows.Add("James", "Brown", "82 Stewart St");
            TestData.Rows.Add("Graham", "Howe", "12 Howard St");
            TestData.Rows.Add("John", "Howe", "78 Short Lane");
            TestData.Rows.Add("Clive", "Smith", "49 Sutherland St");
            TestData.Rows.Add("James", "Owen", "8 Crimson Rd");
            TestData.Rows.Add("Graham", "Brown", "94 Roland St");

            return TestData;
        }

       
    }
}