using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace aspword_cli
{
    class Program
    {
        static void Main(string[] args)
        {

            //示例
            String path = "D://print//docMachine.docx";
            NiceDoc doc = new NiceDoc(path);
            Dictionary<string, object> map = new Dictionary<string, object>();
            map.Add("title", "测试文书记录");
            map.Add("same", 1);
            map.Add("nosame", "无说明");
            map.Add("parson", 6);
            map.Add("prodate", "2019-10-10");
            map.Add("proname", "东莞生产总企业");
            map.Add("isshow", 1);
            doc.setLabel(map);



            List<Dictionary<string, object>> table1 = new List<Dictionary<string, object>>();
            Dictionary<string, object> tableMap1 = new Dictionary<string, object>();
            tableMap1.Add("name", "陈先生");
            tableMap1.Add("date", "2020");
            tableMap1.Add("code", "代码");
            table1.Add(tableMap1);
            Dictionary<string, object> tableMap2 = new Dictionary<string, object>();
            tableMap2.Add("name", "何先生");
            tableMap2.Add("date", "2019");
            tableMap2.Add("code", "代码2");
            table1.Add(tableMap2);
            doc.setTable(table1, "firstTable");

            //doc.save("D://print//docx//" + UUID.randomUUID() + ".docx");
            //doc.saveOnlyComments("D://print//docx//" + UUID.randomUUID() + ".docx");
            doc.savePdf("D://print//docx//C_" + Guid.NewGuid() + ".pdf");

            Console.WriteLine("aspword-cli run!");
        }
    }
}
