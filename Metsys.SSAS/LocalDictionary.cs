using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices;

namespace Metsys.SSAS
{
    public class LocalDictionary
    {
        Dictionary<string,List<string>> localDic = new Dictionary<string, List<string>>();


        private static LocalDictionary self;

        

        private LocalDictionary(Database db)
        {
            DataSourceView dsv = db.DataSourceViews.FindByName("dsv" + db.Name);

            OleDbConnection connection = new OleDbConnection(dsv.DataSource.ConnectionString);
            connection.Open();

            DataSet ds = new DataSet();

            OleDbDataAdapter adapter = new OleDbDataAdapter(
                "select * from dbo." + "MetDictionary",
                connection);

            adapter.Fill(ds);

            DataTable dataTable = ds.Tables[0];

            foreach (DataRow row in dataTable.Rows)
            {
                List<string> list = new List<string>();
                string en = row[1].ToString();
                list.Add(row[2].ToString());
                //list.Add(row[2].ToString());
                localDic.Add(en,list);
            }
        }

        public static LocalDictionary Instance(Database db)
        {
            if (self == null)
            {
                self = new LocalDictionary(db);
            }
            return self;
        }


        public List<string> DoTranslate(string enWord)
        {
            if (localDic.ContainsKey(enWord))
            {
                return localDic[enWord];
            }
            else
            {
                return null;
            }
           
        }


        /// <summary>
        /// translate from english to other language
        /// </summary>
        /// <param name="enWord"></param>
        /// <param name="langId">1.english,2.spanish,3.russian</param>
        /// <returns></returns>
        public string DoTranslate(string enWord, int langId)
        {
            if (langId == 1)
            {
                return enWord;
            }
            else
            {
                if (localDic.ContainsKey(enWord))
                {
                    return localDic[enWord][0];
                }
                else
                {
                    return enWord;
                }    
            }

        }

    }
}
