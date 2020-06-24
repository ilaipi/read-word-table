using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ClosedXML.Excel;
using NPOI.XWPF.UserModel;
using TableIO;
using TableIO.ClosedXml;

namespace read_word_table
{
    class Model
    {
        public string name { get; set; }
        public string sex { get; set; }
        public string minzu { get; set; }
        public string jiguan { get; set; }
        public string birth { get; set; }
        public string zhengzhimm { get; set; }
        public string living { get; set; }
        public string sfzh { get; set; }
        public string phone { get; set; }
        public string xueli { get; set; }
        public string marrige { get; set; }
        public string date { get; set; }
        public string health { get; set; }
        public string company { get; set; }
        public string area { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            string basePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..", "..", "assets");
            var files = Directory.GetFiles(basePath, "*.docx");
            IList<Model> models = new List<Model>();
            string[] fields = new string[] { "name", "sex", "minzu", "jiguan", "birth", "zhengzhimm", "living", "sfzh", "phone", "xueli", "marrige", "date", "health", "company", "area" };
            foreach (var file in files)
            {
                Console.WriteLine("file: " + file);
                models.Add(ParseFile(basePath, file, fields));
            }

            // write models to file.
            using (var workbook = new XLWorkbook(Path.Combine(basePath, "工会会员信息汇总表.xlsx")))
            {
                var worksheet = workbook.Worksheet(1);

                // parameters is (worksheet, startRowNumber, startColumnNuber)
                var tableWriter = new TableFactory().CreateXlsxWriter<Model>(worksheet, 2, 1);
                // parameter is (models, header)
                tableWriter.Write(models, fields);

                workbook.Save();
            }
        }

        private static Model ParseFile(string basePath, string filename, string[] fields)
        {
            int[][] cells = new int[13][] {
                new int[2]{ 0, 1 }, new int[2] { 0, 3 }, new int[2] { 1, 1 },
                new int[2]{ 1, 3 }, new int[2]{ 0, 5 }, new int[2]{ 1, 5 },
                new int[2]{ 2, 1 }, new int[2]{ 6, 5 }, new int[2]{ 7, 5 },
                new int[2]{ 6, 1 }, new int[2]{ 2, 5 }, new int[2]{ 3, 3 },
                new int[2]{ 2, 3 }
            };
            Stream stream = File.OpenRead(filename);
            XWPFDocument doc = new XWPFDocument(stream);
            XWPFTable table = doc.GetTableArray(0);
            Model model = new Model();
            PropertyInfo[] properties = model.GetType().GetProperties();
            for (int i = 0; i < cells.Length; i++)
            {
                int[] c = cells[i];
                XWPFTableRow row = table.GetRow(c[0]);
                XWPFTableCell cell = row.GetCell(c[1]);
                string text = cell.GetText();
                if (fields[i] == "minzu" && !text.EndsWith("族")) // 汉 改为 汉族
                {
                    text = text + "族";
                    cell.SetText(text);
                }
                if (fields[i] == "jiguan" && !text.Contains("省")) // 河南洛阳 改为 河南省洛阳市
                {
                    text = text.Substring(0, 2) + "省" + text.Substring(2, 2) + "市";
                    cell.SetText(text);
                }
                foreach (PropertyInfo t in properties)
                {
                    if (t.Name == fields[i])
                    {
                        t.SetValue(model, text);
                        break;
                    }
                }
            }
            string outputdoc = Path.Combine(basePath, "神经内科", Path.GetFileName(filename));
            doc.Write(File.OpenWrite(outputdoc));
            model.company = "洛阳市第三人民医院";
            model.area = "洛阳市瀍河区";
            stream.Close();
            return model;
        }
    }
}
