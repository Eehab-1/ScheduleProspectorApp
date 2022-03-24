using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduleProspectorApp.MyClasses.XER
{

    public class XerFunctions
    {

        public List<string> TableStringToList(string table)
        {

            List<string> tablelines = new List<string>();

            tablelines = new List<string>(table.Split(new char[] { '|', '\n' }, StringSplitOptions.None));

            return tablelines;

        }


        public int ConvertToNullableInt(string s)
        {
            int num;
            bool check = int.TryParse(s, out num);

            if (check == true)
            {
                return num;
            }
            else
            {
                return 0;
            }



        }




        public double ConvertFromHourtoDays(string _hr_cnt, string day_hr_cnt)
        {

            double lagdouble;

            double day_hr_double;

            bool IsLagParse = double.TryParse(_hr_cnt, out lagdouble);

            bool Is_day_hr_cnt_Parse = double.TryParse(day_hr_cnt, out day_hr_double);

            if ((IsLagParse == true) && (Is_day_hr_cnt_Parse == true))
            {
                double LaginDays = Math.Round((lagdouble / day_hr_double), 2, MidpointRounding.AwayFromZero);

                return LaginDays;
            }

            else
            {
                return 0;
            }




        }



        public DateTime? Convertdatefromstring(string s)
        {

            DateTime xdate;

            bool Isdate = DateTime.TryParse(s, out xdate);

            if (Isdate)
            {
                return xdate;
            }

            else
            {
                return null;
            }



        }




        public void MiniImageDCMA(string ImagePath, string ProjectName, Dictionary<string, bool> IsPassedDic, Dictionary<string, string> PercentDic)

        {







            Bitmap newbmp = new Bitmap(850, 550);

            for (var x = 0; x < newbmp.Width; x++)
            {
                for (var y = 0; y < newbmp.Height; y++)
                {
                    newbmp.SetPixel(x, y, Color.Beige);
                }
            }


            using (Graphics G = Graphics.FromImage(newbmp))




            using (Pen pen = new Pen(Color.Transparent, 8f))
            {

                RectangleF rectf = new RectangleF(315, 5, 250, 50);

                RectangleF bottomtext = new RectangleF(650, 520, 200, 100);

                RectangleF projectname = new RectangleF(5, 520, 200, 100);

                Graphics g = Graphics.FromImage(newbmp);

                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawString("DCMA 14-Point Assessment ", new Font("Tahoma", 12, FontStyle.Bold), Brushes.Chocolate, rectf);

                g.DrawString("https://scheduleprospector.com", new Font("Tahoma", 8), Brushes.Black, bottomtext);

                g.DrawString(ProjectName, new Font("Tahoma", 8), Brushes.Black, projectname);



                pen.LineJoin = System.Drawing.Drawing2D.LineJoin.Round;


                #region Color of box based on IsPassed



                Brush LogicColor;
                Brush LeadColor;
                Brush LagColor;
                Brush FSRelationsColor;
                Brush HardConsColor;
                Brush HighFloatColor;
                Brush NegativeFloatColor;
                Brush HighDurColor;
                Brush InvaliDatesColor;
                Brush ResourcesColor;
                Brush MissedTasksColor;
                Brush CPTColor;
                Brush CPLIColor;
                Brush BEIColor;



                if (IsPassedDic["Logic"])
                {
                    LogicColor = Brushes.ForestGreen;
                }
                else
                {
                    LogicColor = Brushes.Red;
                }


                if (IsPassedDic["Lead"])
                {
                    LeadColor = Brushes.ForestGreen;
                }
                else
                {
                    LeadColor = Brushes.Red;
                }


                if (IsPassedDic["Lag"])
                {
                    LagColor = Brushes.ForestGreen;
                }
                else
                {
                    LagColor = Brushes.Red;
                }


                if (IsPassedDic["FSRelations"])
                {
                    FSRelationsColor = Brushes.ForestGreen;
                }
                else
                {
                    FSRelationsColor = Brushes.Red;
                }

                if (IsPassedDic["HardCons"])
                {
                    HardConsColor = Brushes.ForestGreen;
                }
                else
                {
                    HardConsColor = Brushes.Red;
                }

                if (IsPassedDic["HighFloat"])
                {
                    HighFloatColor = Brushes.ForestGreen;
                }
                else
                {
                    HighFloatColor = Brushes.Red;
                }

                if (IsPassedDic["NegativeFloat"])
                {
                    NegativeFloatColor = Brushes.ForestGreen;
                }
                else
                {
                    NegativeFloatColor = Brushes.Red;
                }

                if (IsPassedDic["HighDur"])
                {
                    HighDurColor = Brushes.ForestGreen;
                }
                else
                {
                    HighDurColor = Brushes.Red;
                }

                if (IsPassedDic["InvaliDates"])
                {
                    InvaliDatesColor = Brushes.ForestGreen;
                }
                else
                {
                    InvaliDatesColor = Brushes.Red;
                }

                if (IsPassedDic["Resources"])
                {
                    ResourcesColor = Brushes.ForestGreen;
                }
                else
                {
                    ResourcesColor = Brushes.Red;
                }

                if (IsPassedDic["MissedTasks"])
                {
                    MissedTasksColor = Brushes.ForestGreen;
                }
                else
                {
                    MissedTasksColor = Brushes.Red;
                }

                if (IsPassedDic["CPT"])
                {
                    CPTColor = Brushes.ForestGreen;
                }
                else
                {
                    CPTColor = Brushes.Red;
                }

                if (IsPassedDic["CPLI"])
                {
                    CPLIColor = Brushes.ForestGreen;
                }
                else
                {
                    CPLIColor = Brushes.Red;
                }


                if (IsPassedDic["BEI"])
                {
                    BEIColor = Brushes.ForestGreen;
                }
                else
                {
                    BEIColor = Brushes.Red;
                }



                #endregion



                G.DrawRectangle(pen, MyRectangle(G, 90, 50, LogicColor, "1. Logic", PercentDic["Logic"], "(should Not exceed 5%)"));
                G.DrawRectangle(pen, MyRectangle(G, 240, 50, LeadColor, "2. Leads", PercentDic["Lead"], "(Leads should Not be used)"));
                G.DrawRectangle(pen, MyRectangle(G, 390, 50, LagColor, "3. Lag", PercentDic["Lag"], "(Should not exceed 5%)"));
                G.DrawRectangle(pen, MyRectangle(G, 540, 50, FSRelationsColor, "4. FS Relationships", PercentDic["FSRelations"], "(Target at least : 90%)"));
                G.DrawRectangle(pen, MyRectangle(G, 690, 50, HardConsColor, "5. Hard Constraints", PercentDic["HardCons"], "(Should not exceed 5%)"));






                G.DrawRectangle(pen, MyRectangle(G, 90, 200, HighFloatColor, "6. High Float", PercentDic["HighFloat"], "(should Not exceed 5%)"));
                G.DrawRectangle(pen, MyRectangle(G, 240, 200, NegativeFloatColor, "7. Negative Float", PercentDic["NegativeFloat"], "(No negative float allowed)"));
                G.DrawRectangle(pen, MyRectangle(G, 390, 200, HighDurColor, "8. HIGH DURATION", PercentDic["HighDur"], "(should Not exceed 5%)"));
                G.DrawRectangle(pen, MyRectangle(G, 540, 200, InvaliDatesColor, "9. Invalid Dates", PercentDic["InvaliDates"], "(No invalid dates allowed)"));
                G.DrawRectangle(pen, MyRectangle(G, 690, 200, ResourcesColor, "10. Resources", PercentDic["Resources"], "(Target 100%)"));


                G.DrawRectangle(pen, MyRectangle(G, 90, 350, MissedTasksColor, "11. Missed Tasks", PercentDic["MissedTasks"], "(Should not exceed 5%)"));
                G.DrawRectangle(pen, MyRectangle(G, 240, 350, CPTColor, "12. Critical Path Test", PercentDic["CPT"], "(No broken logic allowed)"));
                G.DrawRectangle(pen, MyRectangle(G, 390, 350, CPLIColor, "13. Critical Path Length Index", PercentDic["CPLI"], "(No less than 0.95)"));
                G.DrawRectangle(pen, MyRectangle(G, 540, 350, BEIColor, "14 . Baseline Execution Index", PercentDic["BEI"], "(No less than 0.95)"));



            }




            newbmp.Save(ImagePath, ImageFormat.Bmp);
            newbmp.Dispose();




          

        }











        public Rectangle MyRectangle(Graphics g, int p1, int p2, Brush mycolor, string CheckName
            , string CheckPercent, string CheckCondition)
        {
           

            StringFormat format1 = new StringFormat(StringFormatFlags.NoClip);

            format1.LineAlignment = StringAlignment.Center;
            format1.Alignment = StringAlignment.Center;



            Font myfont1 = new Font("Tahoma", 14, FontStyle.Bold);

            Font myfont2 = new Font("Tahoma", 8, FontStyle.Bold);

            Font myfont3 = new Font("Tahoma", 7, FontStyle.Regular);

            StringFormat format2 = new StringFormat(StringFormatFlags.NoClip);

            format2.LineAlignment = StringAlignment.Near;
            format2.Alignment = StringAlignment.Near;

            StringFormat format3 = new StringFormat(StringFormatFlags.NoClip);

            format3.LineAlignment = StringAlignment.Far;
            format3.Alignment = StringAlignment.Far;

            Rectangle r = new Rectangle(p1, p2, 120, 120);

            g.FillRectangle(mycolor, r);



            g.DrawString(CheckPercent, myfont1, Brushes.White, r, format1);

            g.DrawString(CheckName, myfont2, Brushes.White, r, format2);

            g.DrawString(CheckCondition, myfont3, Brushes.White, r, format3);

            return r;

        }





        public ExcelWorksheet ExcelSheetChartFunc(ExcelWorkbook workbook, string SheetName, string ChartText, List<(string, int?)> DataList, string GroupName, string GroupValue)

        {





            var worksheet = workbook.Worksheets.Add(SheetName);





            //Fill the table
            var startCell = worksheet.Cells[1, 1];
            startCell.Offset(0, 0).Value = GroupName;
            startCell.Offset(0, 1).Value = GroupValue;

            for (var i = 0; i < DataList.Count(); i++)
            {
                startCell.Offset(i + 1, 0).Value = DataList[i].Item1;
                startCell.Offset(i + 1, 1).Value = DataList[i].Item2;
            }

            //Add the chart to the sheet
            var pieChart = worksheet.Drawings.AddChart("Chart1", eChartType.Pie);
            pieChart.SetPosition(DataList.Count + 1, 0, 0, 0);
            pieChart.Title.Text = ChartText;
            pieChart.Title.Font.Bold = true;
            pieChart.Title.Font.Size = 12;

            //Set the data range
            var series = pieChart.Series.Add(worksheet.Cells[2, 2, DataList.Count, 2], worksheet.Cells[2, 1, DataList.Count, 1]);
            var pieSeries = (ExcelPieChartSerie)series;
            pieSeries.Explosion = 5;

            //Format the labels
            pieSeries.DataLabel.Font.Bold = true;
            // pieSeries.DataLabel.ShowValue = true;
            pieSeries.DataLabel.ShowPercent = true;
            pieSeries.DataLabel.ShowLeaderLines = true;
            //  pieSeries.DataLabel.Separator = ";";
            pieSeries.DataLabel.Position = eLabelPosition.BestFit;

            //Format the legend
            pieChart.Legend.Add();
            pieChart.Legend.Border.Width = 0;
            pieChart.Legend.Font.Size = 12;
            pieChart.Legend.Font.Bold = true;
            pieChart.Legend.Position = eLegendPosition.Right;

          



            return worksheet;




        }




        public ExcelWorksheet ExcelSheetDCMAFunc(ExcelWorkbook workbook, string SheetName, List<(string, string)> DataList)
        {


            var worksheet = workbook.Worksheets.Add(SheetName);


            //Fill the table
            var startCell = worksheet.Cells[1, 1];
            startCell.Offset(0, 0).Value = "Activity ID";
            startCell.Offset(0, 1).Value = "Acitivty Name";

            for (var i = 0; i < DataList.Count(); i++)
            {
                startCell.Offset(i + 1, 0).Value = DataList[i].Item1;
                startCell.Offset(i + 1, 1).Value = DataList[i].Item2;
            }

            return worksheet;

        }





        public void GlanceAssessment(string folderName, string XerFileContent, XerExtractor xerextractor, XerDCMACheckers xerdcma, XerFunctions xerfunc)
        {


            DataTable TaskDataTable = xerextractor.TASKTable(XerFileContent, xerextractor, xerfunc);

            DataTable TASKPREDTable = xerextractor.TASKPREDTable(XerFileContent, xerextractor, xerfunc);

            var TaskDataTableList = (from t in TaskDataTable.AsEnumerable()

                                     select new
                                     {
                                         task_id = t.Field<string>("task_id"),
                                         task_code = t.Field<string>("task_code"),
                                         task_name = t.Field<string>("task_name"),
                                         total_float = xerfunc.ConvertToNullableInt(t.Field<string>("total_float_hr_cnt")),
                                         status_code = t.Field<string>("status_code"),
                                         task_type = t.Field<string>("task_type"),
                                         cstr_type = t.Field<string>("cstr_type")

                                     }).ToList().OrderBy(a => a.task_code);



            var TASKPREDTableList = (from r in TASKPREDTable.AsEnumerable()
                                     select new
                                     {
                                         pred_type = r.Field<string>("pred_type"),
                                         pred_task_id = r.Field<string>("pred_task_id"),
                                         lag_hr_cnt = xerfunc.ConvertToNullableInt(r.Field<string>("lag_hr_cnt"))

                                     }).ToList();



            DataTable CalendarDataTable = xerextractor.CalendarTable(XerFileContent, xerextractor, xerfunc);

            var TASKPREDDataTable = xerextractor.TASKPREDTable(XerFileContent, xerextractor, xerfunc);

            string ProjectName = xerextractor.ProjectName(XerFileContent, xerextractor, xerfunc);



            // Activity Status Sheet 

            int CompletedCount = TaskDataTableList.Where(a => a.status_code == "TK_Complete").Count();
            int NotStartCount = TaskDataTableList.Where(a => a.status_code == "TK_NotStart").Count();
            int inProgressCount = TaskDataTableList.Where(a => a.status_code == "TK_Active").Count();


            List<(string, int?)> ActsStatusList = new List<(string, int?)>();

            ActsStatusList.Add(("Completed", CompletedCount));
            ActsStatusList.Add(("Not Started", NotStartCount));
            ActsStatusList.Add(("In Progress", inProgressCount));
            ActsStatusList.Add(("", null));


            // Total Float Negative :

            int NegativeFloatCount = TaskDataTableList.Where(a => a.total_float < 0).Count();
            int PositiveFloatCount = TaskDataTableList.Where(a => a.total_float > 0).Count();
            int ZeroFloatCount = TaskDataTableList.Where(a => a.total_float == 0).Count();

            List<(string, int?)> TotalFloatList = new List<(string, int?)>();

            TotalFloatList.Add(("Negative Total Float", NegativeFloatCount));
            TotalFloatList.Add(("Positive Total Float", PositiveFloatCount));
            TotalFloatList.Add(("Total Float equals zero", ZeroFloatCount));
            TotalFloatList.Add(("", null));


            // Task Type 


            int SSCount = TASKPREDTableList.Where(a => a.pred_type == "PR_SS").Count();
            int FFCount = TASKPREDTableList.Where(a => a.pred_type == "PR_FF").Count();
            int SFCount = TASKPREDTableList.Where(a => a.pred_type == "PR_SF").Count();
            int FSCount = TASKPREDTableList.Where(a => a.pred_type == "PR_FS").Count();


            List<(string, int?)> TaskTypeList = new List<(string, int?)>();

            TaskTypeList.Add(("SS_Relationship", SSCount));
            TaskTypeList.Add(("FF_Relationship", FFCount));
            TaskTypeList.Add(("SF_Relationship", SFCount));
            TaskTypeList.Add(("FS_Relationship", FSCount));
            TaskTypeList.Add(("", null));


            // Constraints Type :


            int EmptyCons = TaskDataTableList.Where(a => string.IsNullOrWhiteSpace(a.cstr_type)).Count();

            int HardConsCount = TaskDataTableList.Where(a => a.cstr_type.Contains("MAND")).Count();

            var NonEmpty = TaskDataTableList.Where(a => !string.IsNullOrWhiteSpace(a.cstr_type));

            int SoftConsCount = NonEmpty.Where(a => !a.cstr_type.Contains("MAND")).Count();

            List<(string, int?)> ConsList = new List<(string, int?)>();

            ConsList.Add(("Hard Constraint", HardConsCount));
            ConsList.Add(("Soft Constraint", SoftConsCount));
            ConsList.Add(("No Constraint", HardConsCount));
            ConsList.Add(("", null));


            // Negative Lag (Lead) Positive Lag 


            int NegativeLagCount = TASKPREDTableList.Where(a => a.lag_hr_cnt < 0).Count();
            int PositiveLagCount = TASKPREDTableList.Where(a => a.lag_hr_cnt > 0).Count();
            int NoLagCount = TASKPREDTableList.Where(a => a.lag_hr_cnt == 0).Count();

            List<(string, int?)> LagList = new List<(string, int?)>();

            LagList.Add(("Negative Lag", NegativeLagCount));
            LagList.Add(("Positve Lag", PositiveLagCount));
            LagList.Add(("No Lag", NoLagCount));
            LagList.Add(("", null));






            string datetimenow = DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss");

          

            string pathString = System.IO.Path.Combine(folderName, ProjectName + " Glance  " + datetimenow);
            System.IO.Directory.CreateDirectory(pathString);

            string excelfilepath = pathString + @"\" + "Glance.xlsx";

            var file = new FileInfo(excelfilepath);
            if (file.Exists)
                file.Delete();

            var pck = new ExcelPackage(file);
            var workbook = pck.Workbook;


            xerfunc.ExcelSheetChartFunc(workbook, "Activites_Status", "Status of Activites", ActsStatusList, "Activity_Status", "Status_Count");
            xerfunc.ExcelSheetChartFunc(workbook, "Activites_TotalFloat_Status", "TotalFloat_Status", TotalFloatList, "Activites_TotalFloat", "Activites_Count");
            xerfunc.ExcelSheetChartFunc(workbook, "Relationship_Type", "Relationship_Type", TaskTypeList, "Relationship_Type", "Relationship_Count");
            xerfunc.ExcelSheetChartFunc(workbook, "Constraint_Type", "Constraint_Type", ConsList, "Constraint_Type", "Activites_Count_Per_Contraint");
            xerfunc.ExcelSheetChartFunc(workbook, "Lag_Type", "Lag_Type", LagList, "Lag_Type", "Relationships_Count_Per_Lag_Type");





            pck.Save();










        }


    }





}
