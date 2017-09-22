using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AnalysisServices;

namespace Metsys.SSAS
{
    public class AddScripts
    {
        private Database asDb;
        private LocalDictionary translator;


        public AddScripts(Database db)
        {
            asDb = db;
            translator = LocalDictionary.Instance(db);
        }

        public void AddScript2Dw(string cubeName, int langId)
        {
            cubeName = translator.DoTranslate(cubeName, langId);

            Cube cube = asDb.Cubes.FindByName(cubeName);

            MdxScript script;

            //add mdxscript to cube or use the existing one.
            if (cube.MdxScripts.Count == 0)
            {
                script = new MdxScript();

                script.ID = "MdxScript";
                script.Name = "MdxScript";

                cube.MdxScripts.Add(script);

                Command cmd = new Command();

                script.Commands.Add(cmd);

            }
            else
            {
                script = cube.MdxScripts[0];
            }

            string measures = "Measures";//translator.DoTranslate("Measures", langId); 
            string calcValue = translator.DoTranslate("Measurement Value - Calc", langId); 
            string corrValue = translator.DoTranslate("Measurement Value - Corrected", langId);
            string measuredValue = translator.DoTranslate("Measurement Value - Measured", langId);
            string measureValue = translator.DoTranslate("Measurement Value", langId);
            string mType = translator.DoTranslate("Measurement Type", langId); 
            string measurement = translator.DoTranslate("Measurement", langId); 

            string measurePoint = translator.DoTranslate("Measure Point", langId);
            string aggregation = translator.DoTranslate("Aggregation", langId);
            string avg = translator.DoTranslate("Average", langId);
            string non = translator.DoTranslate("None", langId);
            string datastruct = translator.DoTranslate("Data Structure", langId);
            string prodTime = translator.DoTranslate("Production Time", langId);
            string prodByMonth = translator.DoTranslate("Production by Month", langId);
            string prodByWeek = translator.DoTranslate("Production by Week", langId);
            string prodShift = translator.DoTranslate("Prod Shift", langId);
            string optor = translator.DoTranslate("Operator", langId);
            string scaler = translator.DoTranslate("Scaler Value", langId);

            string prodValue = translator.DoTranslate("Product Value", langId);
            string product = translator.DoTranslate("Product", langId);

            string scriptToRun = "\nCALCULATE;";

            string calcMeasure = "\n\nCREATE MEMBER CURRENTCUBE.[{0}].[{1}]"
                                 + " AS ([{0}].[{2}], [{3}].[{3}].&[3]),"
                                 + " FORMAT_STRING = \"Standard\","
                                 + " NON_EMPTY_BEHAVIOR = {{ [{2}] }},"
                                 + " VISIBLE = 1 ,  ASSOCIATED_MEASURE_GROUP = '{4}';";

            scriptToRun += String.Format(calcMeasure, measures, calcValue, measureValue, mType, measurement);

            string corMeasure = "\n\nCREATE MEMBER CURRENTCUBE.[{0}].[{1}]"
                                 + " AS ([{0}].[{2}], [{3}].[{3}].&[2]),"
                                 + " FORMAT_STRING = \"Standard\","
                                 + " NON_EMPTY_BEHAVIOR = {{ [{2}] }},"
                                 + " VISIBLE = 1 ,  ASSOCIATED_MEASURE_GROUP = '{4}';";


            scriptToRun += String.Format(corMeasure, measures, corrValue, measureValue, mType, measurement);

            string measuredMeasure = "\n\nCREATE MEMBER CURRENTCUBE.[{0}].[{1}]"
                                 + " AS ([{0}].[{2}], [{3}].[{3}].&[1]),"
                                 + " FORMAT_STRING = \"Standard\","
                                 + " NON_EMPTY_BEHAVIOR = {{ [{2}] }},"
                                 + " VISIBLE = 1 ,  ASSOCIATED_MEASURE_GROUP = '{4}';";


            scriptToRun += String.Format(measuredMeasure, measures, measuredValue, measureValue, mType, measurement);

            string avgAggregation = "\n\n--Define Aggregations for Average" +
                                    "\nSCOPE ({0}.[{1}], [{2}].[{3}].&[{4}], [{2}].[{5}].MEMBERS);"
                                    + " this = {0}.[Measurement Value Base] / {0}.[Measurement Count];"
                                    + " END SCOPE;";

            scriptToRun += String.Format(avgAggregation, measures, measureValue, measurePoint, aggregation, avg, datastruct);

            string nonAgg = "\n\n-- Define Aggregations for None.\n" +
                            "SCOPE ([{0}].[{1}].&[{2}], [{0}].[{3}].MEMBERS);";
            scriptToRun += String.Format(nonAgg, measurePoint, aggregation, non, datastruct);
            nonAgg = "\nSCOPE ([{0}].[{1}].[All]);";
            scriptToRun += String.Format(nonAgg, prodTime, prodByMonth);
            nonAgg = "\nSCOPE (DESCENDANTS([{0}].[{1}].CurrentMember, [{0}].[{1}].[{2}], SELF_AND_BEFORE));";
            scriptToRun += String.Format(nonAgg, prodTime, prodByWeek, prodShift);
            nonAgg = "\nThis = ({0}.CurrentMember, [{1}].[{2}].CurrentMember.LastChild);";
            scriptToRun += String.Format(nonAgg, measures, prodTime ,prodByWeek);
            scriptToRun += "\nEND SCOPE;\nEND SCOPE;";
            nonAgg = "\nSCOPE ([{0}].[{1}].[All]);";
            scriptToRun += String.Format(nonAgg, prodTime, prodByWeek);
            nonAgg = "\nSCOPE (DESCENDANTS([{0}].[{1}].CurrentMember, [{0}].[{1}].[{2}], SELF_AND_BEFORE));";
            scriptToRun += String.Format(nonAgg, prodTime, prodByMonth, prodShift);
            nonAgg = "\nThis = ({0}.CurrentMember, [{1}].[{2}].CurrentMember.LastChild);";
            scriptToRun += String.Format(nonAgg, measures, prodTime, prodByMonth);
            scriptToRun += "\nEND SCOPE;\nEND SCOPE;\nEND SCOPE;";

            string typeAgg = "\n\n-- Define Aggregations for Equipment Type Based calcs." +
                             "\nSCOPE ([{0}].[{1}].&[/], [{2}].[{2}].&[3],[{0}].[{3}].MEMBERS);";
            scriptToRun += String.Format(typeAgg, measurePoint, optor, mType, datastruct);
            typeAgg = "\nSCOPE([{0}].[{1}].[All]);";
            scriptToRun += String.Format(typeAgg, prodTime, prodByMonth);
            typeAgg = "\nSCOPE (DESCENDANTS([{0}].[{1}].CurrentMember, [{0}].[{1}].[{2}], SELF_AND_BEFORE));";
            scriptToRun += String.Format(typeAgg, prodTime, prodByWeek, prodShift);
            typeAgg = "\n{0}.[{1}] = {0}.[M1]*StrToValue([{2}].[{3}].CurrentMember.Name)/{0}.[M2];";
            scriptToRun += String.Format(typeAgg, measures, measureValue, measurePoint, scaler);
            scriptToRun += "\nEND SCOPE;\nEND SCOPE;";

            typeAgg = "\nSCOPE([{0}].[{1}].[All]);";
            scriptToRun += String.Format(typeAgg, prodTime, prodByWeek);
            typeAgg = "\nSCOPE (DESCENDANTS([{0}].[{1}].CurrentMember, [{0}].[{1}].[{2}], SELF_AND_BEFORE));";
            scriptToRun += String.Format(typeAgg, prodTime, prodByMonth, prodShift);
            typeAgg = "\n{0}.[{1}] = {0}.[M1]*StrToValue([{2}].[{3}].CurrentMember.Name)/{0}.[M2];";
            scriptToRun += String.Format(typeAgg, measures, measureValue, measurePoint, scaler);
            scriptToRun += "\nEND SCOPE;\nEND SCOPE;\nEND SCOPE;";

            string prodMeasure = "\n\nCREATE MEMBER CURRENTCUBE.[Measures].[{0}]";
            scriptToRun += String.Format(prodMeasure, prodValue);
            prodMeasure = "\nAS Measures.[Product Value Base], \nFORMAT_STRING = \"Standard\", \nNON_EMPTY_BEHAVIOR = { [Product Value Base] },";
            scriptToRun += prodMeasure;
            prodMeasure = "\nVISIBLE = 1 ,  ASSOCIATED_MEASURE_GROUP = '{0}';";
            scriptToRun += string.Format(prodMeasure, product);

            string allFormat = "\n\nSCOPE([Measures].AllMembers); \nFORMAT_STRING (This) = \"#,###.00;(#,###.00)\"; \nFORE_COLOR (This) = iif([Measures].CurrentMember < 0, 255, 16711680); \nEND SCOPE;\n";
            scriptToRun += allFormat;


            script.Commands[0].Text = scriptToRun;
            script.Update();

        }

        public void AddScript2LabDw(string cubeName, int langId)
        {
            cubeName = translator.DoTranslate(cubeName, langId);

            Cube cube = asDb.Cubes.FindByName(cubeName);

            MdxScript script;

            //add mdxscript to cube or use the existing one.
            if (cube.MdxScripts.Count == 0)
            {
                script = new MdxScript();

                script.ID = "MdxScript";
                script.Name = "MdxScript";

                cube.MdxScripts.Add(script);

                Command cmd = new Command();

                script.Commands.Add(cmd);

            }
            else
            {
                script = cube.MdxScripts[0];
            }

            string measures = "Measures";//translator.DoTranslate("Measures", langId); 
            string calcValue = translator.DoTranslate("Measurement Value - Calc", langId);
            string corrValue = translator.DoTranslate("Measurement Value - Corrected", langId);
            string measuredValue = translator.DoTranslate("Measurement Value - Measured", langId);
            string measureValue = translator.DoTranslate("Measurement Value", langId);
            string mType = translator.DoTranslate("Measurement Type", langId);
            string measurement = translator.DoTranslate("Measurement", langId);

            string measurePoint = translator.DoTranslate("Lab Measure Point", langId);
            string aggregation = translator.DoTranslate("Aggregation", langId);
            string avg = translator.DoTranslate("Average", langId);
            string non = translator.DoTranslate("None", langId);
            string datastruct = translator.DoTranslate("Data Structure", langId);

            string scriptToRun = "\nCALCULATE;";

            string calcMeasure = "\n\nCREATE MEMBER CURRENTCUBE.[{0}].[{1}]"
                                 + " AS ([{0}].[{2}], [{3}].[{3}].&[3]),"
                                 + " FORMAT_STRING = \"Standard\","
                                 + " NON_EMPTY_BEHAVIOR = {{ [{2}] }},"
                                 + " VISIBLE = 1 ,  ASSOCIATED_MEASURE_GROUP = '{4}';";

            scriptToRun += String.Format(calcMeasure, measures, calcValue, measureValue, mType, measurement);

            string corMeasure = "\n\nCREATE MEMBER CURRENTCUBE.[{0}].[{1}]"
                                 + " AS ([{0}].[{2}], [{3}].[{3}].&[2]),"
                                 + " FORMAT_STRING = \"Standard\","
                                 + " NON_EMPTY_BEHAVIOR = {{ [{2}] }},"
                                 + " VISIBLE = 1 ,  ASSOCIATED_MEASURE_GROUP = '{4}';";


            scriptToRun += String.Format(corMeasure, measures, corrValue, measureValue, mType, measurement);

            string measuredMeasure = "\n\nCREATE MEMBER CURRENTCUBE.[{0}].[{1}]"
                                 + " AS ([{0}].[{2}], [{3}].[{3}].&[1]),"
                                 + " FORMAT_STRING = \"Standard\","
                                 + " NON_EMPTY_BEHAVIOR = {{ [{2}] }},"
                                 + " VISIBLE = 1 ,  ASSOCIATED_MEASURE_GROUP = '{4}';";


            scriptToRun += String.Format(measuredMeasure, measures, measuredValue, measureValue, mType, measurement);

            string avgAggregation = "\n\n--Define Aggregations for Average" +
                                    "\nSCOPE ({0}.[{1}], [{2}].[{3}].&[{4}], [{2}].[{5}].MEMBERS);"
                                    + " this = {0}.[Measurement Value Base] / {0}.[Measurement Count];"
                                    + " END SCOPE;";
            scriptToRun += String.Format(avgAggregation, measures, measureValue, measurePoint, aggregation, avg, datastruct);

            string allFormat = "\n\nSCOPE([Measures].AllMembers); \nFORMAT_STRING (This) = \"#,###.00;(#,###.00)\"; \nFORE_COLOR (This) = iif([Measures].CurrentMember < 0, 255, 16711680); \nEND SCOPE;\n";
            scriptToRun += allFormat;

            scriptToRun += "/*calculatedmeasures*/ \n";

            script.Commands[0].Text = scriptToRun;
            script.Update();

        }
    }
}
