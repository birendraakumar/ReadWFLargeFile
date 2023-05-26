using System.IO;
using Spire.Doc;

namespace ReadWordFile
{
    class Program
    {
        static void Main(string[] args)
        {
            //Document doc = new Document("C:\\Users\\HP\\Desktop\\document.docx");
            //doc.Range.Replace("sad", "[replaced]", new FindReplaceOptions(FindReplaceDirection.Forward));
            ////doc.Save("C:\\Users\\HP\\Desktop\\Find-And-Replace-Text.docx");
            //doc.Save("C:\\Users\\HP\\Desktop\\Find-And-Replace-Text.pdf", SaveFormat.Pdf);

            //Document doc = new Document();
            //doc.LoadFromFile("C:\\Users\\HP\\Desktop\\document.docx");
            //doc.Replace("passionate", "[replaced]", true, true);
            //doc.SaveToFile("C:\\Users\\HP\\Desktop\\FindandReplace.pdf", FileFormat.PDF);

            Document doc = new Document();
            doc.LoadFromFile("C:\\Users\\HP\\Desktop\\Back_Format.docx");
            doc.Replace("{{@DOCTORNAME}}", "LOUIS C JORDAN M.D.", true, true);
            doc.Replace("{{@NPI}}", "1316976632", true, true);
            doc.Replace("{{@ADDRESS}}", "5716 Cleveland St Suite 200 Virginia Beach VA 23462", true, true);
            doc.Replace("{{@PHONENUMBER}}", "(757) 502-8570", true, true);

            doc.Replace("{{@PATIENT_NAME}}", "CAROLYN HEDRICK", true, true);
            doc.Replace("{{@ADDRESS1}}", "1765 WHITE RIDGE RD", true, true);
            doc.Replace("{{@ADDRESS2}}", "SUTHERLIN VA 24594", true, true);
            doc.Replace("{{@PATIENTPHONE}}", "(434) 822-5034", true, true);
            doc.Replace("{{@PATIENTHEIGHT}}", "5’5", true, true);
            doc.Replace("{{@PATIENTWEIGHT}}", "128", true, true);
            doc.Replace("{{@PATIENTDOB}}", "October 20 1940", true, true);
            doc.Replace("{{@PATENTGENDER}}", "Female", true, true);

            doc.Replace("{{@MEDICARE}}", "9F71F47AE99", true, true);
            doc.Replace("{{@WAISTSIZE}}", "M", true, true);
            doc.Replace("{{@TREATMENTS}}", "REST", true, true);
            doc.Replace("{{@LEVELOFPAIN}}", "8 (Severe)", true, true);
            doc.Replace("{{@EXPERIENCINGTHEPAIN}}", "5 YEARS", true, true);

            doc.Replace("{{@DATE}}", "05-12-2023", true, true);
            doc.Replace("{{@FULLTIMESTAMP}}", "05-12-2023 01:11:22 PM EST IP 98.183.152.34 ]", true, true);

            doc.SaveToFile("C:\\Users\\HP\\Desktop\\FindandReplace.pdf", FileFormat.PDF);
            /**************************************************************/
        }


    }
}
