using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices;

namespace Metsys.SSAS
{
    public class AsDimension
    {

        private static DataTable GetConfigure(Database db)
        {
            // Open a connection to the data source
            DataSourceView dsv = db.DataSourceViews.FindByName("dsv" + db.Name);
            OleDbConnection connection
                = new OleDbConnection(dsv.DataSource.ConnectionString);
            connection.Open();

            DataSet ds = new DataSet();

            OleDbDataAdapter adapter = new OleDbDataAdapter(
                "select * from dbo." + "Mp_Dim_Attr",
                connection);

            adapter.Fill(ds);

            DataTable dataTable = ds.Tables[0];

            return dataTable;
        }

        public static void CreateMpDimension(Database db, int languageId)
        {
            LocalDictionary translator = LocalDictionary.Instance(db);

            string dimName = translator.DoTranslate("Measure Point", languageId);//languageId == 1 ? "Measure Point" : "Measure Point_Se";

            Dimension dim = db.Dimensions.FindByName(dimName);
            if (dim != null)
                dim.Drop();
            dim = db.Dimensions.Add(dimName);

            dim.Type = DimensionType.Regular;
            dim.UnknownMember = UnknownMemberBehavior.Hidden;
            //dim.AttributeAllMemberName = "All Periods";
            dim.Source = new DataSourceViewBinding("dsv" + db.Name);
            dim.StorageMode = DimensionStorageMode.Molap;

            #region Create attributes
            DimensionAttribute attr;
            
            DataTable table = GetConfigure(db);

            foreach (DataRow row in table.Rows )
            {
                string attrId = row[0].ToString();

                string attrName = translator.DoTranslate(row[1].ToString(),languageId);

                string keyCol = row[3].ToString();

                string nameColPr = row[4] == null ? "" : row[4].ToString();
                string nameColSe = row[5] == null ? "" : row[5].ToString();
                string nameCol = languageId == 1 ? nameColPr : nameColSe;

                int usage = int.Parse(row[6].ToString());
                int type = int.Parse(row[7].ToString());
                int orderBy = int.Parse(row[8].ToString());

                string displayEn = row[10] == null ? "" : row[10].ToString();
                string display = translator.DoTranslate(displayEn, languageId);

                string relationId = row[12] == null ? "" : row[12].ToString();
                bool isVisible = bool.Parse(row[13].ToString());
                string tableName = row[14].ToString();


                attr = dim.Attributes.Add(attrId);
                attr.Name = attrName;
                attr.Usage = (AttributeUsage)usage;
                attr.Type = (AttributeType)type;
                attr.OrderBy = (OrderBy)orderBy;
                if (orderBy == 3)
                {
                    string orderByAttrId = row[9].ToString();
                    attr.OrderByAttributeID = orderByAttrId;
                }
                if (display != "" && !string.IsNullOrEmpty(displayEn))
                {
                    attr.AttributeHierarchyDisplayFolder = display;
                }
                attr.AttributeHierarchyVisible = isVisible;
                
                attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, keyCol));
      
                if (nameCol != "")
                {
                    attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, nameCol);
                }

                if (relationId != "")
                {
                    string[] ids = relationId.Split('|');
                    foreach (var id in ids)
                    {
                        attr.AttributeRelationships.Add(new AttributeRelationship(id));
                    }
                }
                
            }


            #endregion

            #region Create hierarchies

            Hierarchy hier;

            hier = dim.Hierarchies.Add(translator.DoTranslate("Data Structure", languageId));
            //hier.AllMemberName = "All Periods";
            hier.Levels.Add(translator.DoTranslate("Plant", languageId)).SourceAttributeID = "Plant";
            hier.Levels.Add(translator.DoTranslate("Plant Area", languageId)).SourceAttributeID = "Plant Area";
            hier.Levels.Add(translator.DoTranslate("Equipment Group", languageId)).SourceAttributeID = "Equipment Group";
            hier.Levels.Add(translator.DoTranslate("Equipment Item", languageId)).SourceAttributeID = "Equipment Item";
            hier.Levels.Add(translator.DoTranslate("Measure Point", languageId)).SourceAttributeID = "Measure Point";
            


            #endregion
            dim.Update();

            //Dimension dim = db.Dimensions.FindByName("Measure Point");
            //if (dim != null)
            //    dim.Drop();
            //dim = db.Dimensions.Add("Measure Point");

            //dim.Type = DimensionType.Regular;
            //dim.UnknownMember = UnknownMemberBehavior.Hidden;
            ////dim.AttributeAllMemberName = "All Periods";
            //dim.Source = new DataSourceViewBinding("dsv" + db.Name);
            //dim.StorageMode = DimensionStorageMode.Molap;

            

            //DimensionAttribute attr;

            //attr = dim.Attributes.Add("Plant");
            //attr.Name = "Plant Test";
            //attr.Usage = AttributeUsage.Regular;
            //attr.Type = AttributeType.Regular;
            //attr.OrderBy = OrderBy.Name;
            //attr.AttributeHierarchyDisplayFolder = "Measure Point";
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "PlantID"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "PlantName");

            //attr = dim.Attributes.Add("Plant Area Number");
            //attr.Name = "Nmuber test";
            //attr.Usage = AttributeUsage.Regular;
            //attr.Type = AttributeType.Regular;
            //attr.OrderBy = OrderBy.Name;
            //attr.AttributeHierarchyVisible = false;
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "PlantAreaNumber"));

            //attr = dim.Attributes.Add("Plant Area");
            //attr.Usage = AttributeUsage.Regular;
            //attr.Type = AttributeType.Regular;
            //attr.OrderBy = OrderBy.AttributeName;
            //attr.OrderByAttributeID = "Plant Area Number";//dim.Attributes.FindByName("Plant Area Number"); //WTF?
            //attr.AttributeHierarchyDisplayFolder = "Measure Point";
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "PlantAreaID"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "PlantAreaName_pr");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Plant"));
            //attr.AttributeRelationships.Add(new AttributeRelationship("Plant Area Number"));

            //attr = dim.Attributes.Add("Equipment Group");
            //attr.Name = "EG_pr";
            //attr.Usage = AttributeUsage.Regular;
            //attr.Type = AttributeType.Regular;
            //attr.OrderBy = OrderBy.Name;
            //attr.AttributeHierarchyDisplayFolder = "Measure Point";
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "EquipmentGroupID"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "EquipmentGroupName_pr");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Plant Area"));

            //attr = dim.Attributes.Add("Equipment Item");
            //attr.Usage = AttributeUsage.Regular;
            //attr.Type = AttributeType.Regular;
            //attr.OrderBy = OrderBy.Name;
            //attr.AttributeHierarchyDisplayFolder = "Measure Point";
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "EquipmentItemID"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "EquipmentItemName_pr");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Equipment Group"));

            //attr = dim.Attributes.Add("Measure Point");
            //attr.Usage = AttributeUsage.Key;
            //attr.Type = AttributeType.Regular;
            //attr.OrderBy = OrderBy.Name;
            //attr.AttributeHierarchyDisplayFolder = "Measure Point";
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "MeasurePointID"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_MeasurePoint", "MeasurePointName_pr");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Equipment Item"));
           
        }

        public static void CreateMpLabDimension(Database db, int languageId)
        {
            LocalDictionary translator = LocalDictionary.Instance(db);

            string dimName = translator.DoTranslate("Lab Measure Point", languageId);

            Dimension dim = db.Dimensions.FindByName(dimName);
            if (dim != null)
                dim.Drop();
            dim = db.Dimensions.Add(dimName);

            dim.Type = DimensionType.Regular;
            dim.UnknownMember = UnknownMemberBehavior.Hidden;
            //dim.AttributeAllMemberName = "All Periods";
            dim.Source = new DataSourceViewBinding("dsv" + db.Name);
            dim.StorageMode = DimensionStorageMode.Molap;

            #region Create attributes
            DimensionAttribute attr;

            DataTable table = GetConfigure(db);

            foreach (DataRow row in table.Rows)
            {
                string attrId = row[0].ToString();
                string attrName = translator.DoTranslate(row[1].ToString(), languageId);

                string keyCol = row[3].ToString();

                string nameColPr = row[4] == null ? "" : row[4].ToString();
                string nameColSe = row[5] == null ? "" : row[5].ToString();
                string nameCol = languageId == 1 ? nameColPr : nameColSe;

                int usage = int.Parse(row[6].ToString());
                int type = int.Parse(row[7].ToString());
                int orderBy = int.Parse(row[8].ToString());

                string displayEn = row[10] == null ? "" : row[10].ToString();
                string display = translator.DoTranslate(displayEn, languageId); ;

                string relationId = row[12] == null ? "" : row[12].ToString();
                bool isVisible = bool.Parse(row[13].ToString());
                string tableName = row[14].ToString();

                


                attr = dim.Attributes.Add(attrId);
                attr.Name = attrName;
                attr.Usage = (AttributeUsage)usage;
                attr.Type = (AttributeType)type;
                attr.OrderBy = (OrderBy)orderBy;
                if (orderBy == 3)
                {
                    string orderByAttrId = row[9].ToString();
                    attr.OrderByAttributeID = orderByAttrId;
                }
                if (display != "" && !string.IsNullOrEmpty(displayEn))
                {
                    attr.AttributeHierarchyDisplayFolder = display;
                }
                attr.AttributeHierarchyVisible = isVisible;
                if (tableName.Contains("Ext"))
                {
                    attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, keyCol));
                }
                else
                {
                    attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName + "_Lab", keyCol));
                }
                if (nameCol != "")
                {
                    attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName + "_Lab", nameCol);
                }

                if (relationId != "")
                {
                    string[] ids = relationId.Split('|');
                    foreach (var id in ids)
                    {
                        attr.AttributeRelationships.Add(new AttributeRelationship(id));
                    }
                }

            }


            #endregion

            #region Create hierarchies

            Hierarchy hier;

            hier = dim.Hierarchies.Add(translator.DoTranslate("Data Structure", languageId));
            //hier.AllMemberName = "All Periods";
            hier.Levels.Add(translator.DoTranslate("Plant", languageId)).SourceAttributeID = "Plant";
            hier.Levels.Add(translator.DoTranslate("Plant Area", languageId)).SourceAttributeID = "Plant Area";
            hier.Levels.Add(translator.DoTranslate("Equipment Group", languageId)).SourceAttributeID = "Equipment Group";
            hier.Levels.Add(translator.DoTranslate("Equipment Item", languageId)).SourceAttributeID = "Equipment Item";
            hier.Levels.Add(translator.DoTranslate("Measure Point", languageId)).SourceAttributeID = "Measure Point";

            #endregion
            dim.Update();

        }

        public static void CreateDateDimension(Database db, int languageId)
        {
            LocalDictionary translator = LocalDictionary.Instance(db);
            string dimName = translator.DoTranslate("Time", languageId);

            Dimension dim = db.Dimensions.FindByName(dimName);
            if (dim != null)
                dim.Drop();
            dim = db.Dimensions.Add(dimName);

            dim.Type = DimensionType.Time;
            dim.UnknownMember = UnknownMemberBehavior.Hidden;
            dim.AttributeAllMemberName = "All Periods";
            dim.Source = new DataSourceViewBinding("dsv" + db.Name);
            dim.StorageMode = DimensionStorageMode.Molap;

            #region Create attributes

            DimensionAttribute attr;

            attr = dim.Attributes.Add("Year");
            attr.Name = translator.DoTranslate("Year", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Years;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Year"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Year_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Year");
            
            attr = dim.Attributes.Add("Quarter");
            attr.Name = translator.DoTranslate("Quarter", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Quarters;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Quarter"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Quarter_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Quarter");
            AttributeRelationship relationship = new AttributeRelationship("Year");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);
            

            attr = dim.Attributes.Add("Month");
            attr.Name = translator.DoTranslate("Month", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Months;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Month"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Month_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Month");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Quarter"));
            relationship = new AttributeRelationship("Quarter");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

            attr = dim.Attributes.Add("Date");
            attr.Name = translator.DoTranslate("Date", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Date;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Date"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Date_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Date");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Month"));
            relationship = new AttributeRelationship("Month");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

            attr = dim.Attributes.Add("Shift");
            attr.Name = translator.DoTranslate("Shift", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.UndefinedTime;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Shift"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Shift_Label");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Shift");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Date"));
            relationship = new AttributeRelationship("Date");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

          

            attr = dim.Attributes.Add("Week Year");
            attr.Name = translator.DoTranslate("Week Year", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Week_Year"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Week_Year_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Week_Year");


            attr = dim.Attributes.Add("Week");
            attr.Name = translator.DoTranslate("Week", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Weeks;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Week"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Week_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Week");
            attr.AttributeRelationships.Add(new AttributeRelationship("Week Year"));

            attr = dim.Attributes.Add("Week Date");
            attr.Name = translator.DoTranslate("Week Date", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Date"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Date_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Date");
            relationship = new AttributeRelationship("Week");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

            attr = dim.Attributes.Add("Week Shift");
            attr.Name = translator.DoTranslate("Week Shift", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.UndefinedTime;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Shift"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Shift_Label");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Shift");
            relationship = new AttributeRelationship("Week Date");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

            attr = dim.Attributes.Add("Week Hour");
            attr.Name = translator.DoTranslate("Week Hour", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Hour"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Hour_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Hour");
            relationship = new AttributeRelationship("Week Shift");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

            attr = dim.Attributes.Add("Daily Status");
            attr.Name = translator.DoTranslate("Daily Status", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "DailyStatus"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "DailyStatusName");

            attr = dim.Attributes.Add("Hour");
            attr.Name = translator.DoTranslate("Hour", languageId);
            attr.Usage = AttributeUsage.Key;
            attr.Type = AttributeType.Hours;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Hour"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Hour_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Time", "Hour");
            attr.AttributeRelationships.Add(new AttributeRelationship("Daily Status"));
            relationship = new AttributeRelationship("Shift");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);
            relationship = new AttributeRelationship("Week Hour");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

            #endregion

            #region Create hierarchies

            Hierarchy hier;

            hier = dim.Hierarchies.Add(translator.DoTranslate("by Month", languageId));
            //hier.AllMemberName = "All Periods";
            hier.Levels.Add(translator.DoTranslate("Year", languageId)).SourceAttributeID = "Year";
            hier.Levels.Add(translator.DoTranslate("Quarter", languageId)).SourceAttributeID = "Quarter";
            hier.Levels.Add(translator.DoTranslate("Month", languageId)).SourceAttributeID = "Month";
            hier.Levels.Add(translator.DoTranslate("Date", languageId)).SourceAttributeID = "Date";
            hier.Levels.Add(translator.DoTranslate("Shift", languageId)).SourceAttributeID = "Shift";
            hier.Levels.Add(translator.DoTranslate("Hour", languageId)).SourceAttributeID = "Hour";

            hier = dim.Hierarchies.Add(translator.DoTranslate("by Week", languageId));
            //hier.AllMemberName = "All Periods";
            hier.Levels.Add(translator.DoTranslate("Week Year", languageId)).SourceAttributeID = "Week Year";
            hier.Levels.Add(translator.DoTranslate("Week", languageId)).SourceAttributeID = "Week";
            hier.Levels.Add(translator.DoTranslate("Date", languageId)).SourceAttributeID = "Week Date";
            hier.Levels.Add(translator.DoTranslate("Shift", languageId)).SourceAttributeID = "Week Shift";
            hier.Levels.Add(translator.DoTranslate("Hour", languageId)).SourceAttributeID = "Week Hour";
            #endregion
            dim.Update();
        }


        /// <summary>
        /// create production time dimension
        /// </summary>
        /// <param name="db"></param>
        /// <param name="languageId"></param>
        public static void CreateProductionDateDimension(Database db, int languageId)
        {
            string tableName = "Dim_ProductionTime";

            LocalDictionary translator = LocalDictionary.Instance(db);
            string dimName = translator.DoTranslate("Production Time", languageId);

            Dimension dim = db.Dimensions.FindByName(dimName);
            if (dim != null)
                dim.Drop();
            dim = db.Dimensions.Add(dimName);

            dim.Type = DimensionType.Time;
            dim.UnknownMember = UnknownMemberBehavior.Hidden;
            dim.AttributeAllMemberName = "All Periods";
            dim.Source = new DataSourceViewBinding("dsv" + db.Name);
            dim.StorageMode = DimensionStorageMode.Molap;

            #region Create attributes

            DimensionAttribute attr;

            attr = dim.Attributes.Add("Year");
            attr.Name = translator.DoTranslate("Prod Year", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Years;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Year"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Year_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Year");

            attr = dim.Attributes.Add("Quarter");
            attr.Name = translator.DoTranslate("Prod Quarter", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Quarters;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Quarter"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Quarter_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Quarter");
            AttributeRelationship relationship = new AttributeRelationship("Year");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);


            attr = dim.Attributes.Add("Month");
            attr.Name = translator.DoTranslate("Prod Month", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Months;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Month"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Month_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Month");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Quarter"));
            relationship = new AttributeRelationship("Quarter");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);


            attr = dim.Attributes.Add("Week Year");
            attr.Name = translator.DoTranslate("Prod Week Year", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "WeekYear"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "WeekYear_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "WeekYear");


            attr = dim.Attributes.Add("Week");
            attr.Name = translator.DoTranslate("Prod Week", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Weeks;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Week"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Week_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Week");
            attr.AttributeRelationships.Add(new AttributeRelationship("Week Year"));

            attr = dim.Attributes.Add("Date");
            attr.Name = translator.DoTranslate("Prod Date", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Date;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Date"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Date_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Date");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Month"));
            relationship = new AttributeRelationship("Month");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

            relationship = new AttributeRelationship("Week");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);


            attr = dim.Attributes.Add("Shift");
            attr.Name = translator.DoTranslate("Prod Shift", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.UndefinedTime;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Shift"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Shift_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Shift");
            //attr.AttributeRelationships.Add(new AttributeRelationship("Date"));
            relationship = new AttributeRelationship("Date");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);





            //attr = dim.Attributes.Add("Week Date");
            //attr.Name = translator.DoTranslate("Week Date", languageId);
            //attr.Usage = AttributeUsage.Regular;
            //attr.Type = AttributeType.Regular;
            //attr.OrderBy = OrderBy.Key;
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Date"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Date_Name");
            //attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Date");
            //relationship = new AttributeRelationship("Week");
            //relationship.RelationshipType = RelationshipType.Rigid;
            //attr.AttributeRelationships.Add(relationship);

            //attr = dim.Attributes.Add("Week Shift");
            //attr.Name = translator.DoTranslate("Week Shift", languageId);
            //attr.Usage = AttributeUsage.Regular;
            //attr.Type = AttributeType.UndefinedTime;
            //attr.OrderBy = OrderBy.Key;
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Shift"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Shift_Name");
            //attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Shift");
            //relationship = new AttributeRelationship("Week Date");
            //relationship.RelationshipType = RelationshipType.Rigid;
            //attr.AttributeRelationships.Add(relationship);

            //attr = dim.Attributes.Add("Week Hour");
            //attr.Name = translator.DoTranslate("Week Hour", languageId);
            //attr.Usage = AttributeUsage.Regular;
            //attr.Type = AttributeType.Regular;
            //attr.OrderBy = OrderBy.Key;
            //attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Hour"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Hour_Name");
            //attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Hour");
            //relationship = new AttributeRelationship("Week Shift");
            //relationship.RelationshipType = RelationshipType.Rigid;
            //attr.AttributeRelationships.Add(relationship);

            attr = dim.Attributes.Add("Daily Status");
            attr.Name = translator.DoTranslate("Daily Status", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "DailyStatus"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "DailyStatusName");

            attr = dim.Attributes.Add("Hour");
            attr.Name = translator.DoTranslate("Prod Hour", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Hours;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "Hour"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Hour_Name");
            attr.ValueColumn = CreateDataItem(db.DataSourceViews[0], tableName, "Hour");
            
            relationship = new AttributeRelationship("Shift");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);


            attr = dim.Attributes.Add("Actual Datetime");
            attr.Name = "Actual Datetime";
            attr.Usage = AttributeUsage.Key;
            attr.Type = AttributeType.Hours;
            attr.OrderBy = OrderBy.Key;
            attr.AttributeHierarchyVisible = false;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], tableName, "ActualDatetime"));

            attr.AttributeRelationships.Add(new AttributeRelationship("Daily Status"));
            relationship = new AttributeRelationship("Hour");
            relationship.RelationshipType = RelationshipType.Rigid;
            attr.AttributeRelationships.Add(relationship);

            //relationship = new AttributeRelationship("Week Hour");
            //relationship.RelationshipType = RelationshipType.Rigid;
            //attr.AttributeRelationships.Add(relationship);


            #endregion

            #region Create hierarchies

            Hierarchy hier;

            hier = dim.Hierarchies.Add(translator.DoTranslate("Production by Month", languageId));
            //hier.AllMemberName = "All Periods";
            hier.Levels.Add(translator.DoTranslate("Prod Year", languageId)).SourceAttributeID = "Year";
            hier.Levels.Add(translator.DoTranslate("Prod Quarter", languageId)).SourceAttributeID = "Quarter";
            hier.Levels.Add(translator.DoTranslate("Prod Month", languageId)).SourceAttributeID = "Month";
            hier.Levels.Add(translator.DoTranslate("Prod Date", languageId)).SourceAttributeID = "Date";
            hier.Levels.Add(translator.DoTranslate("Prod Shift", languageId)).SourceAttributeID = "Shift";
            hier.Levels.Add(translator.DoTranslate("Prod Hour", languageId)).SourceAttributeID = "Hour";

            hier = dim.Hierarchies.Add(translator.DoTranslate("Production by Week", languageId));
            //hier.AllMemberName = "All Periods";
            hier.Levels.Add(translator.DoTranslate("Prod Week Year", languageId)).SourceAttributeID = "Week Year";
            hier.Levels.Add(translator.DoTranslate("Prod Week", languageId)).SourceAttributeID = "Week";
            hier.Levels.Add(translator.DoTranslate("Prod Date", languageId)).SourceAttributeID = "Date";
            hier.Levels.Add(translator.DoTranslate("Prod Shift", languageId)).SourceAttributeID = "Shift";
            hier.Levels.Add(translator.DoTranslate("Prod Hour", languageId)).SourceAttributeID = "Hour";
            #endregion
            dim.Update();
        }

        public static void CreateMeasurementTypeDimension(Database db, int languageId)
        {
            LocalDictionary translator = LocalDictionary.Instance(db);
            string dimName = translator.DoTranslate("Measurement Type", languageId);

            Dimension dim = db.Dimensions.FindByName(dimName);
            if (dim != null)
                dim.Drop();
            dim = db.Dimensions.Add(dimName);

            dim.Type = DimensionType.Regular;
            dim.UnknownMember = UnknownMemberBehavior.None;
            
            dim.Source = new DataSourceViewBinding("dsv" + db.Name);
            dim.StorageMode = DimensionStorageMode.Molap;

            #region Create attributes
            DimensionAttribute attr;

            attr = dim.Attributes.Add("Measurement Type");
            attr.Name = translator.DoTranslate("Measurement Type", languageId);
            attr.Usage = AttributeUsage.Key;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_MeasurementType", "MeasurementTypeID"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_MeasurementType", "MeasurementTypeName");
            
            #endregion

            dim.Update();
        }

        public static void CreateCommentDimension(Database db, int languageId)
        {
            LocalDictionary translator = LocalDictionary.Instance(db);
            string dimName = translator.DoTranslate("Comment", languageId);

            Dimension dim = db.Dimensions.FindByName(dimName);
            if (dim != null)
                dim.Drop();
            dim = db.Dimensions.Add(dimName);

            dim.Type = DimensionType.Regular;
            dim.UnknownMember = UnknownMemberBehavior.None;

            dim.Source = new DataSourceViewBinding("dsv" + db.Name);
            dim.StorageMode = DimensionStorageMode.Molap;

            #region Create attributes
            DimensionAttribute attr;

            attr = dim.Attributes.Add("Comment");
            attr.Name = translator.DoTranslate("Comment", languageId);
            attr.Usage = AttributeUsage.Key;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "CommentId"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "CommentText");

            attr = dim.Attributes.Add("Comment Date");
            attr.Name = translator.DoTranslate("Comment Date", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.AttributeHierarchyVisible = false;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "CommentDate"));

            attr = dim.Attributes.Add("Comment Type");
            attr.Name = translator.DoTranslate("Comment Type", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "CommentTypeId"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "CommentType");

            attr = dim.Attributes.Add("Modified Date");
            attr.Name = translator.DoTranslate("Modified Date", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "ModifiedDate"));

            attr = dim.Attributes.Add("User Name");
            attr.Name = translator.DoTranslate("User Name", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;

            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "UserName"));

            #endregion

            dim.Update();
        }

        public static void CreateProductDimension(Database db, int languageId)
        {
            LocalDictionary translator = LocalDictionary.Instance(db);
            string dimName = translator.DoTranslate("Product", languageId);

            Dimension dim = db.Dimensions.FindByName(dimName);
            if (dim != null)
                dim.Drop();
            dim = db.Dimensions.Add(dimName);

            dim.Type = DimensionType.Regular;
            dim.UnknownMember = UnknownMemberBehavior.None;

            dim.Source = new DataSourceViewBinding("dsv" + db.Name);
            dim.StorageMode = DimensionStorageMode.Molap;

            #region Create attributes
            DimensionAttribute attr;

            attr = dim.Attributes.Add("Aggregation");
            attr.Name = translator.DoTranslate("Aggregation", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_ProductPhenomenon", "Aggregation"));
            //attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "CommentText");

            attr = dim.Attributes.Add("Lot");
            attr.Name = translator.DoTranslate("Lot", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.AttributeHierarchyVisible = false;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_ProductGroup", "ProductGroupId"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_ProductGroup", "Number");
            attr.NameColumn.DataType = OleDbType.WChar;

            attr = dim.Attributes.Add("Batch");
            attr.Name = translator.DoTranslate("Batch", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_ProductItem", "ProductItemId"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_ProductItem", "Number");
            attr.NameColumn.DataType = OleDbType.WChar;
            attr.AttributeRelationships.Add(new AttributeRelationship("Lot"));

            attr = dim.Attributes.Add("Drum");
            attr.Name = translator.DoTranslate("Drum", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.AttributeName;
            attr.OrderByAttributeID = "Product Number";
            attr.AttributeHierarchyVisible = false;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Product", "ProductId"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Product", "ProductNumber");
            attr.NameColumn.DataType = OleDbType.WChar;
            attr.AttributeRelationships.Add(new AttributeRelationship("Batch"));
            attr.AttributeRelationships.Add(new AttributeRelationship("Received"));
            attr.AttributeRelationships.Add(new AttributeRelationship("Product Number"));

            attr = dim.Attributes.Add("Phenomenon ID");
            attr.Name = translator.DoTranslate("Phenomenon", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_ProductPhenomenon", "PhenomenonID"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_ProductPhenomenon", "PhenomenonName_pr");
            attr.AttributeRelationships.Add(new AttributeRelationship("Aggregation"));

            attr = dim.Attributes.Add("Product Number");
            attr.Name = translator.DoTranslate("Product Nmuber", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.AttributeHierarchyVisible = false;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Product", "ProductNumber"));

            attr = dim.Attributes.Add("Product Source Type ID");
            attr.Name = translator.DoTranslate("Product Source Type", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Key;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_ProductSourceType", "ProductSourceTypeId"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_ProductSourceType", "ProductSourceName");

            attr = dim.Attributes.Add("Product Id");
            attr.Name = translator.DoTranslate("ProductKey", languageId);
            attr.Usage = AttributeUsage.Key;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.AttributeHierarchyVisible = false;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Product", "ProductId"));
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Product", "PhenomenonId"));
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Product", "ProductSourceTypeId"));
            attr.NameColumn = CreateDataItem(db.DataSourceViews[0], "Dim_Product", "ProductNumber");
            attr.NameColumn.DataType = OleDbType.WChar;
            attr.AttributeRelationships.Add(new AttributeRelationship("Drum"));
            attr.AttributeRelationships.Add(new AttributeRelationship("Phenomenon ID"));
            attr.AttributeRelationships.Add(new AttributeRelationship("Product Source Type ID"));
            attr.AttributeRelationships.Add(new AttributeRelationship("Units"));

            attr = dim.Attributes.Add("Received");
            attr.Name = translator.DoTranslate("Received", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Product", "Received"));

            attr = dim.Attributes.Add("Units");
            attr.Name = translator.DoTranslate("Units", languageId);
            attr.Usage = AttributeUsage.Regular;
            attr.Type = AttributeType.Regular;
            attr.OrderBy = OrderBy.Name;
            attr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_ProductPhenomenon", "Units"));

            #endregion

            #region Create hierarchies

            Hierarchy hier;

            hier = dim.Hierarchies.Add(translator.DoTranslate("by Lot", languageId));
            hier.Levels.Add(translator.DoTranslate("Lot", languageId)).SourceAttributeID = "Lot";
            hier.Levels.Add(translator.DoTranslate("Batch", languageId)).SourceAttributeID = "Batch";
            hier.Levels.Add(translator.DoTranslate("Drum", languageId)).SourceAttributeID = "Drum";

            #endregion

            dim.Update();
        }

        static DataItem CreateDataItem(DataSourceView dsv, string tableName, string columnName)
        {
            DataTable dataTable = ((DataSourceView)dsv).Schema.Tables[tableName];
            DataColumn dataColumn = dataTable.Columns[columnName];
            return new DataItem(tableName, columnName,
                OleDbTypeConverter.GetRestrictedOleDbType(dataColumn.DataType));
        }

    }
}
