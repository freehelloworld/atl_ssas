using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Microsoft.AnalysisServices;
using System.Data.OleDb;
using System.Globalization;

namespace Metsys.SSAS
{
    class Program
    {
        /// <summary>
        /// used Analysis Management Objects (AMO) to build SSAS project.
        /// data source, data view, cubes(measure groups), dimensions
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {

            #region consts declare

            string serverName = "Data Source=MSDB105\\SQL2016;PROVIDER=MSOLAP;Impersonation Level=Impersonate;";
            string asName = "ALT_AS";
            string dwName = "ALT_DW";
            string databaseServerName = "MSDB105\\SQL2016";

            #endregion

            using (Server svr = new Server())
            {

                

                svr.Connect(serverName);
                Database asDb = svr.Databases.FindByName(asName);
                
                if (asDb == null)
                {
                    asDb = svr.Databases.Add(asName);
                    asDb.DataSourceImpersonationInfo = new ImpersonationInfo(ImpersonationMode.ImpersonateServiceAccount);
                    asDb.Update();
                    AsDataSource.CreateDataSource(asDb, databaseServerName, dwName);

                    AsDatasourceView.CreateDataSourceView(asDb);

                    AsDimension.CreateDateDimension(asDb,1);
                    AsDimension.CreateDateDimension(asDb, 2);

                    AsDimension.CreateProductionDateDimension(asDb, 1);
                    AsDimension.CreateProductionDateDimension(asDb, 2);

                    AsDimension.CreateMpDimension(asDb, 1);
                    AsDimension.CreateMpLabDimension(asDb, 1);

                    AsDimension.CreateMpDimension(asDb, 2);
                    AsDimension.CreateMpLabDimension(asDb, 2);

                    AsDimension.CreateMeasurementTypeDimension(asDb,1);
                    AsDimension.CreateMeasurementTypeDimension(asDb, 2);

                    AsDimension.CreateCommentDimension(asDb,1);
                    AsDimension.CreateCommentDimension(asDb, 2);

                    AsDimension.CreateProductDimension(asDb,1);
                    AsDimension.CreateProductDimension(asDb, 2);

                    AsCube.CreateMeasurePointDwCube(asDb,1);
                    AsCube.CreateMeasurePointDwCube(asDb, 2);

                    AsCube.CreateLabCube(asDb,1);
                    AsCube.CreateLabCube(asDb, 2);
                  
                }
                //AsDatasourceView.CreateDataSourceView(asDb);

                //AsDimension.CreateProductionDateDimension(asDb, 1);
                //AsDimension.CreateProductionDateDimension(asDb, 2);
            }

        }
    }
}
