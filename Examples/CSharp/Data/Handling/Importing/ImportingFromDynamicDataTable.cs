using System.Collections.Generic;
using System.Dynamic;
using System.ComponentModel;
using System.Linq;

namespace Aspose.Cells.Examples.CSharp.Data.Handling.Importing
{
    public class ImportingFromDynamicDataTable
    {
        //Source directory
        static string sourceDir = RunExamples.Get_SourceDirectory();

        //Output directory
        static string outputDir = RunExamples.Get_OutputDirectory();

        public static void Run()
        {
            // ExStart:1
            // Here data is filled into a list of Model class containing two fields only
            var data = GetData();

            // Based upon business some additional fields e.g. unique id is added
            var modifiedData = new Converter().GetModifiedData(data);

            // Modified data is still an object but it is a dynamic one now
            modifiedData.First().Id = 20;

            // Following field is added in the dynamic objects list, but it will not be added to workbook as template file does not have this field
            modifiedData.First().Id2 = 200;


            // Create workbook and fill it with the data
            Workbook workbook = new Workbook(sourceDir + @"ExcelTemplate.xlsx");
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("modifiedData", modifiedData);
            designer.Process();
            designer.Workbook.Save(outputDir + @"ModifiedData.xlsx");

            // Base Model does work but doesn't have the Id
            Workbook workbookRegular = new Workbook(sourceDir + @"ExcelTemplate.xlsx");
            WorkbookDesigner designerRegular = new WorkbookDesigner(workbookRegular);
            designerRegular.SetDataSource("ModifiedData", data);
            designerRegular.Process();
            designerRegular.Workbook.Save(outputDir + @"ModifiedDataRegular.xlsx");

            // ExEnd:1 
        }
        private static List<Model> GetData()
        {
            return new List<Model>
           {
               new Model{ Code = 1 , Name = "One" },
               new Model{ Code = 2 , Name = "Two" },
               new Model{ Code = 3 , Name = "Three" },
               new Model{ Code = 4 , Name = "Four" },
               new Model{ Code = 5 , Name = "Five" }
           };
        }

    }
    public class Model
    {
        public string Name { get; internal set; }
        public int Code { get; internal set; }
    }
    public class Converter
    {
        private int _uniqueNumber;

        public List<dynamic> GetModifiedData(List<Model> data)
        {
            var result = new List<dynamic>();

            result.AddRange(data.ConvertAll<dynamic>(i => AddId(i)));


            return result;
        }

        private dynamic AddId(Model i)
        {
            var result = TransformToDynamic(i);
            result.Id = GetUniqueNumber();
            return result;
        }

        private int GetUniqueNumber()
        {
            var result = _uniqueNumber;
            _uniqueNumber++;
            return result;
        }

        private dynamic TransformToDynamic(object dataObject)
        {

            IDictionary<string, object> expando = new ExpandoObject();

            foreach (PropertyDescriptor property in TypeDescriptor.GetProperties(dataObject.GetType()))
                expando.Add(property.Name, property.GetValue(dataObject));

            return expando as dynamic;

        }
    }

}
