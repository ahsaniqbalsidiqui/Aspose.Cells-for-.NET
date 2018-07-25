using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Cells.Examples.CSharp.Files.Utility
{
    public class EncryptingODSFiles
    {
        //Source directory
        static string sourceDir = RunExamples.Get_SourceDirectory();

        //Output directory
        static string outputDir = RunExamples.Get_OutputDirectory();

        public static void Main()
        {
            //Encrypt an ODS file
            //Encrypted ODS file can only be opened in OpenOffice as Excel does not support encrypted ODS files

            //Initialize loading options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.ODS);

            // Instantiate a Workbook object.
            // Open an ODS file.
            Workbook workbook = new Workbook(sourceDir + "sampleODSFile.ods", loadOptions);

            //Encryption options are not effective for ODS files

            // Password protect the file.
            workbook.Settings.Password = "1234";

            // Save the excel file.
            workbook.Save(outputDir + "outputEncryptedODSFile.ods");

            //Decrypt ODS file
            //Decrypted ODS file can be opened both in Excel and OpenOffice          

            // Set original password
            loadOptions.Password = "1234";

            // Load the encrypted ODS file with the appropriate load options
            Workbook encrypted = new Workbook(sourceDir + "sampleEncryptedODSFile.ods", loadOptions);

            // Unprotect the workbook
            encrypted.Unprotect("1234");

            // Set the password to null
            encrypted.Settings.Password = null;

            // Save the decrypted ODS file
            encrypted.Save(outputDir + "outputDecryptedODSFile.ods");

            Console.WriteLine("Encryption/Decryption of ODS file executed successfully");
        }
    }
}
