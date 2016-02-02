using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Office.Interop.Excel;

namespace UnitTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excel.Workbooks.Open(@"Z:\Users\jahanaf\Documents\Projects\Ssis_Ssrs\Ssis\IssuanceDB\GK IssuanceDB review tool.xlsm");

            var proj = workbook.VBProject;

            //var codeFiles = new List<string>();

            foreach (VBComponent comp in proj.VBComponents)
            {
                var code = comp.CodeModule;

                var fileName = comp.Name + ".txt";
                var content = new StringBuilder();
               
                for (var i = 0; i < code.CountOfLines; i++)
                    content.AppendLine(code.Lines[i + 1, 1]);

                using (TextWriter writer = new StreamWriter(File.Create(Path.Combine(@"Z:\Users\jahanaf\Documents\Projects\Ssis_Ssrs\Ssis\IssuanceDB\Vba", fileName))))
                {
                    writer.Write(content);
                }
            }

            workbook.Close();
        }
    }
}
