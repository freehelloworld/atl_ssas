
using Microsoft.AnalysisServices;

namespace Metsys.SSAS
{
    public class AsDataSource
    {
        public static DataSource CreateDataSource(Database asDb, string databaseServerName, string dwName)
        {
            string datasourceName = "ds" + asDb.Name;
            // Create the data source
            DataSource ds = asDb.DataSources.Add(datasourceName);
            ds.ConnectionString = "Provider = SQLNCLI11.1; Data Source = " + databaseServerName + "; Integrated Security = SSPI; Initial Catalog = " + dwName;

            // Send the data source definition to the server.
            ds.Update();

            return ds;
        }

    }
}
