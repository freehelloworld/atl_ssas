using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices;

namespace Metsys.SSAS
{
    public class AsCube
    {
        public static void CreateMeasurePointDwCube(Database db, int languageId)
        {

            LocalDictionary translator = LocalDictionary.Instance(db);

            string cubeName = translator.DoTranslate("Measure Point DW", languageId);

            Cube cube = db.Cubes.FindByName(cubeName);
            if (cube != null)
                cube.Drop();

            cube = db.Cubes.Add(cubeName);
           
            cube.Source = new DataSourceViewBinding("dsv" + db.Name);
            cube.StorageMode = StorageMode.Molap;

            #region Create cube dimensions
            Dimension dim;
            //use production time instead of time.
            //dim = db.Dimensions.GetByName(translator.DoTranslate("Time", languageId));
            dim = db.Dimensions.GetByName(translator.DoTranslate("Production Time", languageId));
            cube.Dimensions.Add(dim.ID);

            dim = db.Dimensions.GetByName(translator.DoTranslate("Measurement Type", languageId));
            cube.Dimensions.Add(dim.ID);

            dim = db.Dimensions.GetByName(translator.DoTranslate("Measure Point", languageId));
            cube.Dimensions.Add(dim.ID);

            dim = db.Dimensions.GetByName(translator.DoTranslate("Comment", languageId));
            cube.Dimensions.Add(dim.ID);

            dim = db.Dimensions.GetByName(translator.DoTranslate("Product", languageId));
            cube.Dimensions.Add(dim.ID);
            #endregion
            
            #region Create measure groups

            CreateMeasurementMeasureGroup(cube, languageId);

            CreateProductMeasureGroup(cube, languageId);
            
            CreateCommentGroup(cube, languageId);

            #endregion

            cube.Update(UpdateOptions.ExpandFull);  
        }

        static void CreateMeasurementMeasureGroup(Cube cube, int languageId)
        {
            Database db = cube.Parent;
            LocalDictionary translator = LocalDictionary.Instance(db);

            string mgName = translator.DoTranslate("Measurement", languageId);

            MeasureGroup mg = cube.MeasureGroups.FindByName(mgName);
            if (mg != null)
                mg.Drop();
            mg = cube.MeasureGroups.Add("Measurement");
            mg.Name = mgName;
            mg.StorageMode = StorageMode.Molap;
            mg.ProcessingMode = ProcessingMode.Regular;
            //mg.Type = MeasureGroupType.Sales;  

            #region Create measures

            Measure meas;
            
            meas = mg.Measures.Add(translator.DoTranslate("Measurement Value", languageId));
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.FormatString = "Standard";
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "MeasurementValue");

            meas = mg.Measures.Add("X1");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "X1");

            meas = mg.Measures.Add("X2");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "X2");

            meas = mg.Measures.Add("Y1");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "Y1");

            meas = mg.Measures.Add("Y2");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "Y2");

            meas = mg.Measures.Add("Measurement Value Base");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "MeasurementValue");

            meas = mg.Measures.Add("Measurement Count");
            meas.AggregateFunction = AggregationFunction.Count;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "MeasurementValue");

            meas = mg.Measures.Add("M1");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "M1");

            meas = mg.Measures.Add("M2");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "M2");

            #endregion

            #region Create measure group dimensions

            CubeDimension cubeDim;
            RegularMeasureGroupDimension regMgDim;
            MeasureGroupAttribute mgAttr;

            //   Mapping dimension and key column from fact table  
            //      > select dimension and add it to the measure group  
            //cubeDim = cube.Dimensions.GetByName("Time");
            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Production Time", languageId));
            regMgDim = new RegularMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(regMgDim);

            //      > add key column from dimension and map it with   
            //        the surrogate key in the fact table  
            mgAttr = regMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName("Actual Datetime").ID);   // this is dimension key column  
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "MeasurementDateTime"));   // this surrogate key in fact table  

            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Measurement Type", languageId));
            regMgDim = new RegularMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(regMgDim);
            mgAttr = regMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName(translator.DoTranslate("Measurement Type", languageId)).ID);
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "MeasurementTypeID"));

            //cubeDim = cube.Dimensions.GetByName("Measure Point");
            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Measure Point", languageId));
            regMgDim = new RegularMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(regMgDim);
            mgAttr = regMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName(translator.DoTranslate("Measure Point", languageId)).ID);
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_Measurement", "MeasurePointID"));

            #endregion

            #region Create partitions

            CreateMeasurementGroupPartitions(mg, db, languageId);

            #endregion
        }
       

        static void CreateProductMeasureGroup(Cube cube, int languageId)
        {

            Database db = cube.Parent;
            LocalDictionary translator = LocalDictionary.Instance(db);

            MeasureGroup mg = cube.MeasureGroups.FindByName(translator.DoTranslate("Product", languageId));
            if (mg != null)
                mg.Drop();
            mg = cube.MeasureGroups.Add("Product");
            mg.Name = translator.DoTranslate("Product", languageId);
            mg.StorageMode = StorageMode.Molap;
            mg.ProcessingMode = ProcessingMode.Regular;
            //mg.Type = MeasureGroupType.Sales;  

            #region Create measures

            Measure meas;

            meas = mg.Measures.Add("Product Value Base");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_ProductValue", "Value");

            meas = mg.Measures.Add("Product Count");
            meas.AggregateFunction = AggregationFunction.Count;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_ProductValue", "Value");
            

            #endregion

            #region Create measure group dimensions

            CubeDimension cubeDim;
            RegularMeasureGroupDimension regMgDim;
            MeasureGroupAttribute mgAttr;

            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Production Time", languageId));
            regMgDim = new RegularMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(regMgDim);

            mgAttr = regMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName("Actual Datetime").ID);   // this is dimension key column  
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_ProductValue", "Produced"));   // this surrogate key in fact table  

            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Product", languageId));
            regMgDim = new RegularMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(regMgDim);
            mgAttr = regMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName(translator.DoTranslate("ProductKey", languageId)).ID);
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_ProductValue", "ProductId"));
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_ProductValue", "PhenomenonId"));
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_ProductValue", "ProductSourceTypeId"));


            #endregion

            #region Create partitions

            CreateProductGroupPartitions(mg, db);

            #endregion
        }


        static void CreateCommentGroup(Cube cube, int languageId)
        {
            Database db = cube.Parent;
            LocalDictionary translator = LocalDictionary.Instance(db);

            string mgName = translator.DoTranslate("Comment", languageId);

            MeasureGroup mg = cube.MeasureGroups.FindByName(mgName);
            if (mg != null)
                mg.Drop();
            mg = cube.MeasureGroups.Add(mgName);
            mg.StorageMode = StorageMode.Molap;
            mg.ProcessingMode = ProcessingMode.Regular;
            //mg.Type = MeasureGroupType.Sales;  

            #region Create measures

            Measure meas;

            meas = mg.Measures.Add(translator.DoTranslate("Comment Count", languageId));
            meas.AggregateFunction = AggregationFunction.Count;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "CommentId");

            #endregion

            #region Create measure group dimensions

            CubeDimension cubeDim;
            ReferenceMeasureGroupDimension refMgDim;
            DegenerateMeasureGroupDimension degMgDim;
            MeasureGroupAttribute mgAttr;


            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Comment", languageId));
            degMgDim = new DegenerateMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(degMgDim);
            mgAttr = degMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName(translator.DoTranslate("Comment", languageId)).ID);
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Dim_Comment", "CommentId"));

            
            //this relys on the previous relationship
            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Production Time", languageId));

            CubeDimension refDim = cube.Dimensions.GetByName(translator.DoTranslate("Comment", languageId));

            refMgDim = new ReferenceMeasureGroupDimension();

            refMgDim.Materialization = ReferenceDimensionMaterialization.Regular;

            refMgDim.CubeDimensionID = cubeDim.ID;

            refMgDim.IntermediateCubeDimensionID = refDim.ID;

            refMgDim.IntermediateGranularityAttributeID = refDim.Dimension.Attributes.GetByName(translator.DoTranslate("Comment Date", languageId)).ID;

            mg.Dimensions.Add(refMgDim);

            mgAttr = refMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName("Actual Datetime").ID);   // this is dimension key column  
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
             
            
            #endregion

            #region Create partitions

            //CreateMeasurementGroupPartitions(mg, db);

            #endregion
        }

        static DataItem CreateDataItem(DataSourceView dsv, string tableName, string columnName)
        {
            DataTable dataTable = ((DataSourceView)dsv).Schema.Tables[tableName];
            DataColumn dataColumn = dataTable.Columns[columnName];
            return new DataItem(tableName, columnName,
                OleDbTypeConverter.GetRestrictedOleDbType(dataColumn.DataType));
        }

        static void CreateMeasurementGroupPartitions(MeasureGroup mg, Database db, int languageId)
        {
            LocalDictionary translator = LocalDictionary.Instance(db);

            Partition part;
            part = mg.Partitions.FindByName("EquipGroup_1");
            if (part != null)
                part.Drop();
            part = mg.Partitions.Add("EquipGroup_1");
            part.StorageMode = StorageMode.Molap;
            string query = "SELECT [dbo].[Fact_Measurement].[MeasurementDateTime],[dbo].[Fact_Measurement].[MeasurePointID],[dbo].[Fact_Measurement].[MeasurementTypeID],[dbo].[Fact_Measurement].[MeasurementValue],[dbo].[Fact_Measurement].[X1],[dbo].[Fact_Measurement].[X2],[dbo].[Fact_Measurement].[Y1],[dbo].[Fact_Measurement].[Y2],[dbo].[Fact_Measurement].[EquipmentGroupID],[dbo].[Fact_Measurement].[M1],[dbo].[Fact_Measurement].[M2]" +
		                    " FROM [dbo].[Fact_Measurement]" +
		                    " WHERE [EquipmentGroupID] = 1";

            part.Source = new QueryBinding(db.DataSources[0].ID, query);
            string slice = "[{0}].[{1}].&[1]";
            //part.Slice = "[Measure Point].[Equipment Group].&[100]";
            part.Slice = string.Format(slice, translator.DoTranslate("Measure Point",languageId), translator.DoTranslate("Equipment Group",languageId));
        }

        static void CreateProductGroupPartitions(MeasureGroup mg, Database db)
        {
            Partition part;
            part = mg.Partitions.FindByName("Fact_ProductValue");
            if (part != null)
                part.Drop();
            part = mg.Partitions.Add("Fact_ProductValue");
            part.StorageMode = StorageMode.Molap;

            part.Source = new TableBinding(db.DataSources[0].ID, "dbo", "Fact_ProductValue");
           
        }
        

        public static void CreateLabCube(Database db, int languageId)
        {
            LocalDictionary translator = LocalDictionary.Instance(db);

            string cubeName = translator.DoTranslate("Lab DW", languageId);

            // Create the Adventure Works cube  
            Cube cube = db.Cubes.FindByName(cubeName);
            if (cube != null)
                cube.Drop();

            cube = db.Cubes.Add(cubeName);

            cube.Source = new DataSourceViewBinding("dsv" + db.Name);
            cube.StorageMode = StorageMode.Molap;

            #region Create cube dimensions
            Dimension dim;
            dim = db.Dimensions.GetByName(translator.DoTranslate("Production Time", languageId));
            cube.Dimensions.Add(dim.ID);

            dim = db.Dimensions.GetByName(translator.DoTranslate("Measurement Type", languageId));
            cube.Dimensions.Add(dim.ID);

            dim = db.Dimensions.GetByName(translator.DoTranslate("Lab Measure Point", languageId));
            cube.Dimensions.Add(dim.ID);

            #endregion

            #region Create measure groups

            CreateLabMeasureGroup(cube, languageId);

            #endregion

            cube.Update(UpdateOptions.ExpandFull);
        }

        static void CreateLabMeasureGroup(Cube cube, int languageId)
        {
            Database db = cube.Parent;
            LocalDictionary translator = LocalDictionary.Instance(db);

            MeasureGroup mg = cube.MeasureGroups.FindByName("Measurement");
            if (mg != null)
                mg.Drop();
            mg = cube.MeasureGroups.Add("Measurement");
            mg.Name = translator.DoTranslate("Measurement", languageId);
            mg.StorageMode = StorageMode.Molap;
            mg.ProcessingMode = ProcessingMode.Regular;
            //mg.Type = MeasureGroupType.Sales;  

            #region Create measures

            Measure meas;

            meas = mg.Measures.Add(translator.DoTranslate("Measurement Value", languageId));
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.FormatString = "Standard";
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "MeasurementValue");

            meas = mg.Measures.Add("X1");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "X1");

            meas = mg.Measures.Add("X2");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "X2");

            meas = mg.Measures.Add("Y1");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "Y1");

            meas = mg.Measures.Add("Y2");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "Y2");

            meas = mg.Measures.Add("Measurement Value Base");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "MeasurementValue");

            meas = mg.Measures.Add("Measurement Count");
            meas.AggregateFunction = AggregationFunction.Count;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "MeasurementValue");

            meas = mg.Measures.Add("M1");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "M1");

            meas = mg.Measures.Add("M2");
            meas.AggregateFunction = AggregationFunction.Sum;
            meas.Visible = false;
            meas.Source = CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "M2");

            #endregion

            #region Create measure group dimensions

            CubeDimension cubeDim;
            RegularMeasureGroupDimension regMgDim;
            MeasureGroupAttribute mgAttr;

            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Production Time", languageId));
            regMgDim = new RegularMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(regMgDim);

            //      > add key column from dimension and map it with   
            //        the surrogate key in the fact table  
            mgAttr = regMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName("Actual Datetime").ID);   // this is dimension key column  
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "MeasurementDateTime"));   // this surrogate key in fact table  

            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Measurement Type", languageId));
            regMgDim = new RegularMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(regMgDim);
            mgAttr = regMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName(translator.DoTranslate("Measurement Type", languageId)).ID);
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "MeasurementTypeID"));

            cubeDim = cube.Dimensions.GetByName(translator.DoTranslate("Lab Measure Point", languageId));
            regMgDim = new RegularMeasureGroupDimension(cubeDim.ID);
            mg.Dimensions.Add(regMgDim);
            mgAttr = regMgDim.Attributes.Add(cubeDim.Dimension.Attributes.GetByName(translator.DoTranslate("Measure Point",languageId)).ID);
            mgAttr.Type = MeasureGroupAttributeType.Granularity;
            mgAttr.KeyColumns.Add(CreateDataItem(db.DataSourceViews[0], "Fact_Measurement_Lab", "MeasurePointID"));

            #endregion

            #region Create partitions

            CreateLabMeasureGroupPartitions(mg, db);

            #endregion
        }

        static void CreateLabMeasureGroupPartitions(MeasureGroup mg, Database db)
        {
            Partition part;
            part = mg.Partitions.FindByName("Fact_Measurement_Lab");
            if (part != null)
                part.Drop();
            part = mg.Partitions.Add("Fact_Measurement_Lab");
            part.StorageMode = StorageMode.Molap;

            part.Source = new TableBinding(db.DataSources[0].ID, "dbo", "Fact_Measurement_Lab");

        }
       

    }
}
