using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices;

namespace Metsys.SSAS
{
    public class AsDatasourceView
    {
        public static void CreateDataSourceView(Database db)
        {
            // Find the data source view
            DataSourceView dsv = db.DataSourceViews.FindByName("dsv" + db.Name);
            if (dsv != null)
                dsv.Drop();

            string dsvName = "dsv" + db.Name;
            // Create the data source view
            dsv = db.DataSourceViews.Add(dsvName);
            dsv.DataSourceID = "ds" + db.Name;
            dsv.Schema = new DataSet();
            dsv.Schema.Locale = CultureInfo.CurrentCulture;

           
            // Open a connection to the data source
            OleDbConnection connection
                = new OleDbConnection(dsv.DataSource.ConnectionString);
            connection.Open();
            
            // Add the date dimension table
            AddView(dsv, connection, "Fact_Measurement");

            AddView(dsv, connection, "Fact_Measurement_Lab");

            AddView(dsv, connection, "Dim_MeasurePoint");

            AddView(dsv, connection, "Dim_MeasurePoint_Lab");

            AddView(dsv, connection, "Dim_MeasurementType");

            AddView(dsv, connection, "Dim_MeasurePoint_Ext");

            AddView(dsv, connection, "Dim_ProductionTime");

            AddNamedQuery(dsv, connection, "Dim_Time");

            AddView(dsv, connection, "Dim_Comment");

            AddView(dsv, connection, "Fact_ProductValue");

            AddView(dsv, connection, "Dim_Product");

            AddTable(dsv, connection, "Dim_ProductGroup");

            AddTable(dsv, connection, "Dim_ProductItem");

            AddView(dsv, connection, "Dim_ProductPhenomenon");

            AddTable(dsv, connection, "Dim_ProductSourceType");


            AddRelation(dsv, "Fact_Measurement", "MeasurePointID", "Dim_MeasurePoint", "MeasurePointID");

            AddRelation(dsv, "Fact_Measurement", "MeasurementTypeID", "Dim_MeasurementType", "MeasurementTypeID");

            AddRelation(dsv, "Fact_Measurement", "MeasurementDateTime", "Dim_Time", "DateTime");

            AddRelation(dsv, "Fact_Measurement", "MeasurementDateTime", "Dim_ProductionTime", "ActualDatetime");

            AddRelation(dsv, "Dim_MeasurePoint", "MeasurePointID", "Dim_MeasurePoint_Ext", "MeasurePointID");

            
            AddRelation(dsv, "Fact_Measurement_Lab", "MeasurePointID", "Dim_MeasurePoint_Lab", "MeasurePointID");

            AddRelation(dsv, "Fact_Measurement_Lab", "MeasurementTypeID", "Dim_MeasurementType", "MeasurementTypeID");

            AddRelation(dsv, "Fact_Measurement_Lab", "MeasurementDateTime", "Dim_Time", "DateTime");

            AddRelation(dsv, "Fact_Measurement_Lab", "MeasurementDateTime", "Dim_ProductionTime", "ActualDatetime");

            AddRelation(dsv, "Dim_MeasurePoint_Lab", "MeasurePointID", "Dim_MeasurePoint_Ext", "MeasurePointID");


            AddRelation(dsv, "Dim_Comment", "CommentDate", "Dim_Time", "DateTime");

            AddRelation(dsv, "Dim_Comment", "CommentDate", "Dim_ProductionTime", "ActualDatetime");


            AddRelation(dsv, "Fact_ProductValue", "Produced", "Dim_Time", "DateTime");

            AddRelation(dsv, "Fact_ProductValue", "Produced", "Dim_ProductionTime", "ActualDatetime");

            AddRelation(dsv, "Fact_ProductValue", "ProductId", "Dim_Product", "ProductId");

            AddRelation(dsv, "Dim_Product", "ProductItemId", "Dim_ProductItem", "ProductItemId");

            AddRelation(dsv, "Dim_Product", "PhenomenonID", "Dim_ProductPhenomenon", "PhenomenonID");

            AddRelation(dsv, "Dim_Product", "ProductSourceTypeId", "Dim_ProductSourceType", "ProductSourceTypeId");

            AddRelation(dsv, "Dim_ProductItem", "ProductGroupId", "Dim_ProductGroup", "ProductGroupId");

            // Add the fact table
            dsv.Update(); // Send the data source view definition to the server
        }

        public static void AddView(DataSourceView dsv, OleDbConnection connection, string tableName)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter(
                "select * from dbo." + tableName,
                connection);
            DataTable[] dataTables = adapter.FillSchema(dsv.Schema,
                SchemaType.Mapped, tableName);
            DataTable dataTable = dataTables[0];
            if (dataTable.ExtendedProperties.Count == 0)
            {
                dataTable.ExtendedProperties.Add("TableType", "View");
                dataTable.ExtendedProperties.Add("DbSchemaName", "dbo");
                dataTable.ExtendedProperties.Add("DbTableName", tableName);
                dataTable.ExtendedProperties.Add("FriendlyName", tableName);
            }
        }

        public static void AddTable(DataSourceView dsv, OleDbConnection connection, string tableName)
        {
            OleDbDataAdapter adapter = new OleDbDataAdapter(
                "select * from DataMart." + tableName,
                connection);
            DataTable[] dataTables = adapter.FillSchema(dsv.Schema,
                SchemaType.Mapped, tableName);
            DataTable dataTable = dataTables[0];
            if (dataTable.ExtendedProperties.Count == 0)
            {
                dataTable.ExtendedProperties.Add("TableType", "Table");
                dataTable.ExtendedProperties.Add("DbSchemaName", "DataMart");
                dataTable.ExtendedProperties.Add("DbTableName", tableName);
                dataTable.ExtendedProperties.Add("FriendlyName", tableName);
            }
            return;
        }

        static void AddRelation(DataSourceView dsv, String fkTableName, String fkColumnName, String pkTableName, String pkColumnName)
        {
            DataColumn fkColumn
                = dsv.Schema.Tables[fkTableName].Columns[fkColumnName];
            DataColumn pkColumn
                = dsv.Schema.Tables[pkTableName].Columns[pkColumnName];
            dsv.Schema.Relations.Add("FK_" + fkTableName + "_"
                + fkColumnName + pkTableName, pkColumn, fkColumn, true);
        }

        static DataItem CreateDataItem(DataSourceView dsv, string tableName, string columnName)
        {
            DataTable dataTable = ((DataSourceView)dsv).Schema.Tables[tableName];
            DataColumn dataColumn = dataTable.Columns[columnName];
            return new DataItem(tableName, columnName,
                OleDbTypeConverter.GetRestrictedOleDbType(dataColumn.DataType));
        }

        private static bool AddNamedQuery(DataSourceView dsv, OleDbConnection connection, string groupName)
        {
            bool isAdded = false; // Used to determine if the named query already exists
            string pivotStatement = "SELECT  t.DateTime, t.Hour, t.Hour_Name, t.Shift, t.Shift_Name, t.Shift_Label, t.Shift_Sort, t.Date, t.Date_Name, t.Week, t.Week_Name, t.Month, t.Month_Name, t.Quarter, t.Quarter_Name, t.Year, t.Year_Name, "+
                         " t.Week_Year, t.Week_Year_Name, ISNULL(ds.Status, 1) AS DailyStatus, CASE isnull(ds.Status, 1)" +
                         " WHEN 1 THEN 'Pending' WHEN 2 THEN 'Interim' WHEN 3 THEN 'Awaiting' WHEN 4 THEN 'Approved' END AS DailyStatusName "+
                         " FROM    dbo.Dim_Time AS t LEFT OUTER JOIN " +
                         " DataMart.FactDailyDataStatus AS ds ON t.Date = ds.DataStartDate";
            OleDbDataAdapter adapter = new OleDbDataAdapter(
                pivotStatement,
                connection);
            DataTable[] dataTables = adapter.FillSchema(dsv.Schema,
                SchemaType.Mapped, groupName);

            DataTable dataTable = dataTables[0];

            if (dataTable.ExtendedProperties.Count == 0)
            {
                isAdded = true;
                dataTable.ExtendedProperties.Add("TableType", "View");
                dataTable.ExtendedProperties["IsLogical"] = "True";
            }
            dataTable.ExtendedProperties["QueryDefinition"] = pivotStatement;
            return isAdded;
        }

    }
}
