using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduleProspectorApp.MyClasses.XER
{


    public class XerDCMACheckers
    {


        //1. Logic

      

        public (List<(string, string)> NonLogicActs, string LogicPercent, bool IsPassed)
       MissingLogicCheck(DataTable TaskDataTable, DataTable RelationsDataTable)

        {

            var TaskDataTableList = from t in TaskDataTable.AsEnumerable()
                                    where t.Field<string>("status_code") != "TK_Complete"
                                    select new
                                    {
                                        task_id = t.Field<string>("task_id"),
                                        task_code = t.Field<string>("task_code"),
                                        task_name = t.Field<string>("task_name")
                                    };




            var RelationsDataTableList = from r in RelationsDataTable.AsEnumerable()
                                         select new
                                         {
                                             task_id = r.Field<string>("task_id"),
                                             pred_task_id = r.Field<string>("pred_task_id")
                                         };

            var foundSucList = from t in TaskDataTableList
                               join r in RelationsDataTableList
                               on t.task_id equals r.task_id
                               select t;

            var foundPrecList = from t in TaskDataTableList
                                join r in RelationsDataTableList
                                on t.task_id equals r.pred_task_id
                                select t;


            var NonfoundSuc = TaskDataTableList.Except(foundSucList);

            var NonfoundPrec = TaskDataTableList.Except(foundPrecList);

            NonfoundPrec.ToList().AddRange(NonfoundSuc);

            NonfoundPrec.Distinct();

            double NonfoundPrecCount = NonfoundPrec.Count();
            double TaskDataTableListCount = TaskDataTableList.Count();

            double LogicPercent = (NonfoundPrecCount / TaskDataTableListCount) * 100;

            double LogicPercentRounded = Math.Round(LogicPercent, 2);

            string LogicPercentRoundedText = LogicPercentRounded.ToString() + "%";

            bool ispassed = true;

            if (LogicPercentRounded > 5)
            {
                ispassed = false;
            }

            List<(string, string)> NonLogicActs = new List<(string, string)>();


            foreach (var item in NonfoundPrec)
            {
                NonLogicActs.Add((item.task_code, item.task_name));
            }


            return (NonLogicActs, LogicPercentRoundedText, ispassed);



        }


        //2. Leads 


        public (List<(string, string)> PreActivitesWithLead, string LeadPercent, bool IsPassed)
    LeadsCheck(DataTable TaskDataTable, DataTable RelationsDataTable, DataTable CalendarDataTable)

        {

            XerFunctions xerfunctions = new XerFunctions();


            var CalendarDataTableList = from c in CalendarDataTable.AsEnumerable()
                                        select new
                                        {
                                            clndr_id = c.Field<string>("clndr_id"),
                                            clndr_name = c.Field<string>("clndr_name"),
                                            day_hr_cnt = c.Field<string>("day_hr_cnt")
                                        };


            var TaskDataTableList = from t in TaskDataTable.AsEnumerable()
                                    where t.Field<string>("status_code") != "TK_Complete"
                                    select new
                                    {
                                        task_id = t.Field<string>("task_id"),
                                        task_code = t.Field<string>("task_code"),
                                        task_name = t.Field<string>("task_name"),
                                        clndr_id = t.Field<string>("clndr_id")
                                    };




            var RelationsDataTableList = from r in RelationsDataTable.AsEnumerable()
                                         select new
                                         {
                                             task_id = r.Field<string>("task_id"),
                                             pred_task_id = r.Field<string>("pred_task_id"),
                                             lag_hr_cnt = r.Field<string>("lag_hr_cnt")

                                         };


            var NonCompletedTasksRelationList = (from t in TaskDataTableList
                                                 from r in RelationsDataTableList
                                                 where t.task_id == r.task_id || t.task_id == r.pred_task_id
                                                 select new
                                                 {

                                                     Lag = xerfunctions.ConvertToNullableInt(r.lag_hr_cnt),
                                                     Task_Id = r.task_id,
                                                     Pred_Task_Id = r.pred_task_id
                                                 });


            var NegativeLagList = NonCompletedTasksRelationList.ToList().Distinct().Where(a => a.Lag < 0);

         




            // Acts have Leads with thier Predecs

            var ActsWithLeadsWithSuc = from t in TaskDataTableList
                                       from r in NegativeLagList
                                       where t.task_id == r.Pred_Task_Id
                                       select new
                                       {
                                           clndr_id = t.clndr_id,
                                           task_id = t.task_id,
                                           task_code = t.task_code,
                                           task_name = t.task_name,
                                           lag_hr_cnt = r.Lag


                                       };




            var ActsWithLeadsWithPrecDays = from a in ActsWithLeadsWithSuc
                                            from c in CalendarDataTableList
                                            where a.clndr_id == c.clndr_id
                                            select new
                                            {
                                                clndr_id = a.clndr_id,
                                                task_id = a.task_id,
                                                task_code = a.task_code,
                                                task_name = a.task_name,
                                                lag_hr_cnt = a.lag_hr_cnt,
                                                day_hr_cnt = c.day_hr_cnt,

                                                LaginDays = xerfunctions.ConvertFromHourtoDays(a.lag_hr_cnt.ToString(), c.day_hr_cnt)

                                            };



            var ActsWithLeadsWithPrecDaysSorted = ActsWithLeadsWithPrecDays.OrderBy(a => a.task_code);


            double ActsWithLeadsWithPrecDaysSortedCount = ActsWithLeadsWithPrecDaysSorted.Count();
            double TaskDataTableListCount = TaskDataTableList.Count();

            double LeadPercent = Math.Round((ActsWithLeadsWithPrecDaysSortedCount / TaskDataTableListCount) * 100, 2);

            string LeadPercentText = LeadPercent.ToString() + "%";

            bool ispassed = true;

            if (LeadPercent != 0)
            {
                ispassed = false;
            }


            List<(string, string)> PreActivitesWithLeadList = new List<(string, string)>();

            foreach (var item in ActsWithLeadsWithSuc)
            {
                PreActivitesWithLeadList.Add((item.task_code, item.task_name));
            }

            return (PreActivitesWithLeadList, LeadPercentText, ispassed);


        }



        //3. Lag 


        public (List<(string, string)> PreActivitesWithLag, string LagPercent, bool IsPassed)
 LagCheck(DataTable TaskDataTable, DataTable RelationsDataTable, DataTable CalendarDataTable)

        {

            XerFunctions xerfunctions = new XerFunctions();


            var CalendarDataTableList = from c in CalendarDataTable.AsEnumerable()
                                        select new
                                        {
                                            clndr_id = c.Field<string>("clndr_id"),
                                            clndr_name = c.Field<string>("clndr_name"),
                                            day_hr_cnt = c.Field<string>("day_hr_cnt")
                                        };


            var TaskDataTableList = from t in TaskDataTable.AsEnumerable()
                                    where t.Field<string>("status_code") != "TK_Complete"
                                    select new
                                    {
                                        task_id = t.Field<string>("task_id"),
                                        task_code = t.Field<string>("task_code"),
                                        task_name = t.Field<string>("task_name"),
                                        clndr_id = t.Field<string>("clndr_id")
                                    };




            var RelationsDataTableList = from r in RelationsDataTable.AsEnumerable()
                                         select new
                                         {
                                             task_id = r.Field<string>("task_id"),
                                             pred_task_id = r.Field<string>("pred_task_id"),
                                             lag_hr_cnt = r.Field<string>("lag_hr_cnt")

                                         };


            var NonCompletedTasksRelationList = (from t in TaskDataTableList
                                                 from r in RelationsDataTableList
                                                 where t.task_id == r.task_id || t.task_id == r.pred_task_id
                                                 select new
                                                 {

                                                     Lag = xerfunctions.ConvertToNullableInt(r.lag_hr_cnt),
                                                     Task_Id = r.task_id,
                                                     Pred_Task_Id = r.pred_task_id
                                                 });


            var PositiveeLagList = NonCompletedTasksRelationList.ToList().Distinct().Where(a => a.Lag > 0);






            // Acts have Leads with thier Predecs

            var ActsWithLagsWithSuc = from t in TaskDataTableList
                                      from r in PositiveeLagList
                                      where t.task_id == r.Pred_Task_Id
                                      select new
                                      {
                                          clndr_id = t.clndr_id,
                                          task_id = t.task_id,
                                          task_code = t.task_code,
                                          task_name = t.task_name,
                                          lag_hr_cnt = r.Lag


                                      };




            var ActsWithLagWithPrecDays = from a in ActsWithLagsWithSuc
                                          from c in CalendarDataTableList
                                          where a.clndr_id == c.clndr_id
                                          select new
                                          {
                                              clndr_id = a.clndr_id,
                                              task_id = a.task_id,
                                              task_code = a.task_code,
                                              task_name = a.task_name,
                                              lag_hr_cnt = a.lag_hr_cnt,
                                              day_hr_cnt = c.day_hr_cnt,

                                              LaginDays = xerfunctions.ConvertFromHourtoDays(a.lag_hr_cnt.ToString(), c.day_hr_cnt)

                                          };



            var ActsWithLagWithPrecDaysSorted = ActsWithLagWithPrecDays.OrderBy(a => a.task_code);


            double ActsWithLagWithPrecDaysSortedCount = ActsWithLagWithPrecDaysSorted.Count();
            double TaskDataTableListCount = TaskDataTableList.Count();

            double LagPercent = Math.Round((ActsWithLagWithPrecDaysSortedCount / TaskDataTableListCount) * 100, 2);

            string LagPercentText = LagPercent.ToString() + "%";

            bool ispassed = true;

            if (LagPercent > 5)
            {
                ispassed = false;
            }


            List<(string, string)> PreActivitesWithLagList = new List<(string, string)>();

            foreach (var item in ActsWithLagsWithSuc)
            {
                PreActivitesWithLagList.Add((item.task_code, item.task_name));
            }

            return (PreActivitesWithLagList, LagPercentText, ispassed);


        }






        //4. FS Relationship
        public (List<(string, string)> NonFSActivites, string FSPercent, bool IsPassed)
        FSRelationshipCheck(DataTable TaskDataTable, DataTable RelationsDataTable)

        {

            var TaskList = from t in TaskDataTable.AsEnumerable()
                           select new
                           {

                               task_code = t.Field<string>("task_code"),
                               task_id = t.Field<string>("task_id"),
                               task_name = t.Field<string>("task_name")

                           };


            var RelationsList = from r in RelationsDataTable.AsEnumerable()
                                select new
                                {
                                    pred_type = r.Field<string>("pred_type"),
                                    pred_task_id = r.Field<string>("pred_task_id")
                                };





            double PR_FS_Count = RelationsList.Where(r => r.pred_type == "PR_FS").Count();

            double RelationListCount = RelationsList.Count();

            double PR_FS_Percent = (PR_FS_Count / RelationListCount) * 100;

            double PR_FS_Percent_Rounded = Math.Round(PR_FS_Percent, 2);

            string PR_FS_Percent_Rounded_Text = PR_FS_Percent_Rounded.ToString() + "%";



            List<(string, string)> NonFSActivites = new List<(string, string)>();


            var NonFSActivitesSelected = from r in RelationsList
                                         join t in TaskList
                                         on r.pred_task_id equals t.task_id
                                         where r.pred_type != "PR_FS"
                                         select t;

            foreach (var item in NonFSActivitesSelected)
            {
                NonFSActivites.Add((item.task_code, item.task_name));
            }


            bool ispassed = true;

            if (PR_FS_Percent_Rounded < 90)
            {
                ispassed = false;
            }


            return (NonFSActivites, PR_FS_Percent_Rounded_Text, ispassed);




        }




        //5. HARD CONSTRAINTS


        public (List<(string, string)> HardConsActs, string HardConsPercent, bool IsPassed)
            HardConsCheck(DataTable TaskDataTable)

        {

            var TaskDataTableList = from t in TaskDataTable.AsEnumerable()
                                    where t.Field<string>("status_code") != "TK_Complete"
                                    select new
                                    {
                                        task_id = t.Field<string>("task_id"),
                                        task_code = t.Field<string>("task_code"),
                                        task_name = t.Field<string>("task_name"),
                                        cstr_type = t.Field<string>("cstr_type")
                                    };


            var TaskDataTableListHardCons = TaskDataTableList.Where(a => a.cstr_type.Contains("MAND"));

            List<(string, string)> HardConsActsList = new List<(string, string)>();

            foreach (var item in TaskDataTableListHardCons)
            {
                HardConsActsList.Add((item.task_code, item.task_name));
            }


            double TaskDataTableListCount = TaskDataTableList.Count();
            double TaskDataTableListHardConsCount = TaskDataTableListHardCons.Count();

            double HardConsPercent = Math.Round(((TaskDataTableListHardConsCount / TaskDataTableListCount) * 100), 2);

            string HardConsPercentText = HardConsPercent.ToString() + "%";

            bool ispassed = true;

            if (HardConsPercent > 5)
            {
                ispassed = false;
            }


            return (HardConsActsList, HardConsPercentText, ispassed);


        }



        //6. HIGH FLOAT

        public (List<(string, string)> HighFloatActs, string HighFloatPercent, bool ispassed)
            HighFloatCheck(DataTable TaskDataTable, DataTable CalendarDataTable, XerFunctions xerfunc)

        {




            var CalendarDataTableList = from c in CalendarDataTable.AsEnumerable()
                                        select new
                                        {
                                            clndr_id = c.Field<string>("clndr_id"),
                                            clndr_name = c.Field<string>("clndr_name"),
                                            day_hr_cnt = c.Field<string>("day_hr_cnt")
                                        };


            var TaskDataTableList = from t in TaskDataTable.AsEnumerable()
                                    where t.Field<string>("status_code") != "TK_Complete"
                                    select new
                                    {
                                        task_id = t.Field<string>("task_id"),
                                        task_code = t.Field<string>("task_code"),
                                        task_name = t.Field<string>("task_name"),
                                        clndr_id = t.Field<string>("clndr_id"),
                                        total_float_hr_cnt = t.Field<string>("total_float_hr_cnt")
                                    };



            var TaskwithFloatinDays = from t in TaskDataTableList
                                      from c in CalendarDataTableList
                                      where t.clndr_id == c.clndr_id
                                      select new
                                      {
                                          t.task_id,
                                          t.task_code,
                                          t.task_name,
                                          t.total_float_hr_cnt,
                                          day_hr_cnt = c.day_hr_cnt,
                                          TotalFloatDays = xerfunc.ConvertFromHourtoDays(t.total_float_hr_cnt, c.day_hr_cnt)
                                      };



            var TaskwithFloatinDays44More = TaskwithFloatinDays.Where(a => a.TotalFloatDays > 44)
                .ToList().OrderBy(b => b.task_code);

            double TaskDataTableListCount = TaskDataTableList.Count();

            double TaskwithFloatinDays44MoreCount = TaskwithFloatinDays44More.Count();

            double TF44ExtraPercent = (TaskwithFloatinDays44MoreCount / TaskDataTableListCount) * 100;

            double TF44ExtraPercentRounded = Math.Round(TF44ExtraPercent, 2, MidpointRounding.AwayFromZero);

            string TF44ExtraPercentRoundedText = TF44ExtraPercentRounded.ToString() + "%";

            bool ispassed = true;

            if (TF44ExtraPercentRounded > 5)
            {
                ispassed = false;
            }


            List<(string, string)> ActsHighFloat = new List<(string, string)>();

            foreach (var item in TaskwithFloatinDays44More)
            {
                ActsHighFloat.Add((item.task_code, item.task_name));
            }



            return (ActsHighFloat, TF44ExtraPercentRoundedText, ispassed);


        }




        //7. Negative Float 

        public (List<(string, string)> NegativeFloatActs, string NegativeFloatPercent, bool ispassed)
           NegativeFloatCheck(DataTable TaskDataTable, DataTable CalendarDataTable, XerFunctions xerfunc)
        {


            var CalendarDataTableList = from c in CalendarDataTable.AsEnumerable()
                                        select new
                                        {
                                            clndr_id = c.Field<string>("clndr_id"),
                                            clndr_name = c.Field<string>("clndr_name"),
                                            day_hr_cnt = c.Field<string>("day_hr_cnt")
                                        };


            var TaskDataTableList = from t in TaskDataTable.AsEnumerable()
                                    where t.Field<string>("status_code") != "TK_Complete"
                                    select new
                                    {
                                        task_id = t.Field<string>("task_id"),
                                        task_code = t.Field<string>("task_code"),
                                        task_name = t.Field<string>("task_name"),
                                        clndr_id = t.Field<string>("clndr_id"),
                                        total_float_hr_cnt = t.Field<string>("total_float_hr_cnt")
                                    };



            var TaskwithFloatinDays = from t in TaskDataTableList
                                      from c in CalendarDataTableList
                                      where t.clndr_id == c.clndr_id
                                      select new
                                      {
                                          t.task_id,
                                          t.task_code,
                                          t.task_name,
                                          t.total_float_hr_cnt,
                                          day_hr_cnt = c.day_hr_cnt,
                                          TotalFloatDays = xerfunc.ConvertFromHourtoDays(t.total_float_hr_cnt, c.day_hr_cnt)
                                      };



            var TaskwithFloatinDaysNegative = TaskwithFloatinDays.Where(a => a.TotalFloatDays < 0)
                .ToList().OrderBy(b => b.task_code);

            double TaskDataTableListCount = TaskDataTableList.Count();

            double TaskwithFloatinDaysNegativeCount = TaskwithFloatinDaysNegative.Count();

            double TF44ExtraPercent = (TaskwithFloatinDaysNegativeCount / TaskDataTableListCount) * 100;

            double TF44ExtraPercentRounded = Math.Round(TF44ExtraPercent, 2, MidpointRounding.AwayFromZero);

            string TF44ExtraPercentRoundedText = TF44ExtraPercentRounded.ToString() + "%";

            bool ispassed = true;

            if (TF44ExtraPercentRounded > 0)
            {
                ispassed = false;
            }


            List<(string, string)> ActsNegativeFloat = new List<(string, string)>();

            foreach (var item in TaskwithFloatinDaysNegative)
            {
                ActsNegativeFloat.Add((item.task_code, item.task_name));
            }



            return (ActsNegativeFloat, TF44ExtraPercentRoundedText, ispassed);





        }



        // 8.  HIGH DURATION

        public (List<(string, string)> HighDurationActs, string HighDurationPercent, bool ispassed)
           HighDurationCheck(DataTable TaskDataTable, DataTable CalendarDataTable, XerFunctions xerfunc)

        {

            var CalendarDataTableList = from c in CalendarDataTable.AsEnumerable()
                                        select new
                                        {
                                            clndr_id = c.Field<string>("clndr_id"),
                                            clndr_name = c.Field<string>("clndr_name"),
                                            day_hr_cnt = c.Field<string>("day_hr_cnt")
                                        };


            var TaskDataTableList = from t in TaskDataTable.AsEnumerable()
                                    where t.Field<string>("status_code") != "TK_Complete"
                                    select new
                                    {
                                        task_id = t.Field<string>("task_id"),
                                        task_code = t.Field<string>("task_code"),
                                        task_name = t.Field<string>("task_name"),
                                        clndr_id = t.Field<string>("clndr_id"),
                                        target_drtn_hr_cnt = t.Field<string>("target_drtn_hr_cnt")
                                    };



            var TaskwithDurationinDays = (from t in TaskDataTableList
                                          from c in CalendarDataTableList
                                          where t.clndr_id == c.clndr_id
                                          select new
                                          {
                                              t.task_id,
                                              t.task_code,
                                              t.task_name,
                                              t.target_drtn_hr_cnt,
                                              day_hr_cnt = c.day_hr_cnt,
                                              DurationinDays = xerfunc.ConvertFromHourtoDays(t.target_drtn_hr_cnt, c.day_hr_cnt)
                                          }).ToList().OrderBy(a => a.task_code);







            var TaskwithDurationinDaysmore44 = TaskwithDurationinDays.Where(a => a.DurationinDays > 44);


            List<(string, string)> ActsDurationmore44 = new List<(string, string)>();

            foreach (var item in TaskwithDurationinDaysmore44)
            {
                ActsDurationmore44.Add((item.task_code, item.task_name));
            }

            double TaskwithDurationinDaysmore44Count = TaskwithDurationinDaysmore44.Count();

            double TaskDataTableListCount = TaskDataTableList.Count();

            double Durmore44Percent = Math.Round(((TaskwithDurationinDaysmore44Count / TaskDataTableListCount) * 100), 2, MidpointRounding.AwayFromZero);

            string Durmore44PercentText = Durmore44Percent.ToString() + "%";

            bool ispassed = true;

            if (Durmore44Percent > 5)
            {
                ispassed = false;
            }


            return (ActsDurationmore44, Durmore44PercentText, ispassed);

        }



        //9. INVALID DATES


        public (List<(string, string)> InvalidDatesActs, string InvalidDatesPercent, bool ispassed)
           InvalidDatesCheck(DataTable TaskDataTable, DateTime? date_date, XerFunctions xerfunc)
        {

            var TaskDataTableList = (from t in TaskDataTable.AsEnumerable()
                                     where t.Field<string>("status_code") != "TK_Complete"
                                     select new
                                     {
                                         act_start_date = t.Field<string>("act_start_date"),
                                         act_end_date = t.Field<string>("act_end_date"),
                                         target_start_date = t.Field<string>("target_start_date"),
                                         target_end_date = t.Field<string>("target_end_date"),
                                         task_code = t.Field<string>("task_code"),
                                         task_name = t.Field<string>("task_name")

                                     }).ToList().OrderBy(a => a.task_code);


            var TaskDataTableListConverted = from t in TaskDataTableList
                                             select new
                                             {
                                                 act_start_date = xerfunc.Convertdatefromstring(t.act_start_date),
                                                 act_end_date = xerfunc.Convertdatefromstring(t.act_end_date),
                                                 target_start_date = xerfunc.Convertdatefromstring(t.target_start_date),
                                                 target_end_date = xerfunc.Convertdatefromstring(t.target_end_date),
                                                 task_code = t.task_code,
                                                 task_name = t.task_name
                                             };




            var InvalidDatesActsList = from t in TaskDataTableListConverted
                                       where t.target_start_date < date_date ||
                                       t.target_end_date < date_date ||
                                       t.act_start_date > date_date ||
                                       t.act_end_date > date_date
                                       select t;



            List<(string, string)> InvalidDatesActs = new List<(string, string)>();

            foreach (var item in InvalidDatesActsList)
            {
                InvalidDatesActs.Add((item.task_code, item.task_name));
            }

            double InvalidDatesActsListCount = InvalidDatesActsList.Count();

            double TaskDataTableListCount = TaskDataTableList.Count();

            double InvalidPercent = Math.Round(((InvalidDatesActsListCount / TaskDataTableListCount) * 100), 2, MidpointRounding.AwayFromZero);

            string InvalidPercentText = InvalidPercent.ToString() + "%";

            bool ispassed = true;

            if (InvalidDatesActsListCount > 0)
            {
                ispassed = false;
            }


            return (InvalidDatesActs, InvalidPercentText, ispassed);


        }







        //10. RESOURCES


        public (List<(string, string)> ActswithMissedResources, string MissedResourcesPercent, bool ispassed)
           MissedResourcesCheck(DataTable TaskDataTable)

        {

            var TaskDataTableList = (from t in TaskDataTable.AsEnumerable()
                                     where t.Field<string>("status_code") != "TK_Complete"
                                     select new
                                     {
                                         target_drtn_hr_cnt = t.Field<string>("target_drtn_hr_cnt"),
                                         task_code = t.Field<string>("task_code"),
                                         task_name = t.Field<string>("task_name"),
                                         rsrc_id = t.Field<string>("rsrc_id")

                                     }).ToList().OrderBy(a => a.task_code);


            var TaskDataTableListDurmorezero = TaskDataTableList.Where(a => a.target_drtn_hr_cnt != "0");

            var TaskDataTableListDurmorezerowithMissedResources = TaskDataTableListDurmorezero.Where(a => string.IsNullOrWhiteSpace(a.rsrc_id));


            List<(string, string)> ActswithMissedResources = new List<(string, string)>();


            foreach (var item in TaskDataTableListDurmorezerowithMissedResources)
            {
                ActswithMissedResources.Add((item.task_code, item.task_name));
            }


            double ActswithMissedResourcesCount = ActswithMissedResources.Count;

            double TaskDataTableListDurmorezeroCount = TaskDataTableListDurmorezero.Count();

            double MissedResourcesPercent =
       Math.Round((ActswithMissedResourcesCount / TaskDataTableListDurmorezeroCount) * 100, 2, MidpointRounding.AwayFromZero);

            string MissedResourcesPercentText = MissedResourcesPercent.ToString() + "%";

            bool ispassed = true;

            if (ActswithMissedResourcesCount > 0)
            {
                ispassed = false;
            }


            return (ActswithMissedResources, MissedResourcesPercentText, ispassed);

        }



        //11. MISSED TASKS


        public (List<(string, string)> MissedList, string MissedTaskPercent, bool ispassed)

            MissedTasksCheck(DataTable TaskDataTable, DateTime? DataDate, XerFunctions xerfunc)


        {



            var TaskDataTableList = (from t in TaskDataTable.AsEnumerable()

                                     select new
                                     {
                                         act_start_date = t.Field<string>("act_start_date"),
                                         act_end_date = t.Field<string>("act_end_date"),
                                         target_start_date = t.Field<string>("target_start_date"),
                                         target_end_date = t.Field<string>("target_end_date"),
                                         task_code = t.Field<string>("task_code"),
                                         task_name = t.Field<string>("task_name")

                                     }).ToList().OrderBy(a => a.task_code);


            var TaskDataTableListConverted = from t in TaskDataTableList
                                             select new
                                             {
                                                 act_start_date = xerfunc.Convertdatefromstring(t.act_start_date),
                                                 act_end_date = xerfunc.Convertdatefromstring(t.act_end_date),
                                                 target_start_date = xerfunc.Convertdatefromstring(t.target_start_date),
                                                 target_end_date = xerfunc.Convertdatefromstring(t.target_end_date),
                                                 task_code = t.task_code,
                                                 task_name = t.task_name
                                             };







            var TaskDataTableListConvertedcheck1 = TaskDataTableListConverted.Where(a => a.target_end_date <= DataDate);

            var TaskDataTableListConvertedcheck2 = TaskDataTableListConvertedcheck1
                .Where(a => (a.target_end_date - a.act_end_date).Value.Ticks < 0);


            List<(string, string)> MissedList = new List<(string, string)>();

            foreach (var item in TaskDataTableListConvertedcheck2)
            {
                MissedList.Add((item.task_code, item.task_name));
            }


            double TaskDataTableListConvertedcheck2Count = TaskDataTableListConvertedcheck2.Count();

            double TaskDataTableListConvertedcheck1Count = TaskDataTableListConvertedcheck1.Count();

            double MissedTasksPercent;

            if ((TaskDataTableListConvertedcheck2Count == 0) && (TaskDataTableListConvertedcheck1Count == 0))
            {
                MissedTasksPercent = 0;
            }
            else
            {


                MissedTasksPercent = (TaskDataTableListConvertedcheck2Count / TaskDataTableListConvertedcheck1Count) * 100;
            }


            double MissedTasksPercentRounded = Math.Round(MissedTasksPercent, 2, MidpointRounding.AwayFromZero);

            string MissedTasksPercentText = MissedTasksPercentRounded.ToString() + "%";

            bool ispassed = true;

            if (MissedTasksPercentRounded > 5)
            {
                ispassed = false;
            }


            return (MissedList, MissedTasksPercentText, ispassed);

        }





        //12. CRITICAL PATH TEST



        public (List<(string, string)> ActsWithoutList, string ActsPercent, bool ispassed)

            CPTCheck(DataTable TaskDataTable, DataTable RelationsDataTable)
        {

            var TaskDataTableList = (from t in TaskDataTable.AsEnumerable()

                                     select new
                                     {

                                         task_code = t.Field<string>("task_code"),
                                         task_name = t.Field<string>("task_name"),
                                         task_id = t.Field<string>("task_id")

                                     }).ToList().OrderBy(a => a.task_code);


            var RelationsDataTableList = (from r in RelationsDataTable.AsEnumerable()
                                          select new
                                          {
                                              task_id = r.Field<string>("task_id"),
                                              pred_task_id = r.Field<string>("pred_task_id"),


                                          }).ToList();


            var ActswithSucc = from t in TaskDataTableList
                               join r in RelationsDataTableList
                               on t.task_id equals r.pred_task_id
                               select t;

            var ActsWithoutSucc = TaskDataTableList.Except(ActswithSucc).ToList();

            var ActswithPrec = from t in TaskDataTableList
                               join r in RelationsDataTableList
                               on t.task_id equals r.task_id
                               select t;

            var ActsWithoutPrec = TaskDataTableList.Except(ActswithPrec).ToList(); ;




            List<(string, string)> TotalActsWithoutSuccORPred = new List<(string, string)>();

            foreach (var item in ActsWithoutSucc)
            {
                TotalActsWithoutSuccORPred.Add((item.task_code, item.task_name));
            }

            foreach (var item in ActsWithoutPrec)
            {
                TotalActsWithoutSuccORPred.Add((item.task_code, item.task_name));
            }

            double ActsWithoutSuccCount = ActsWithoutSucc.Count();

            double ActsWithoutPrecCount = ActsWithoutPrec.Count();

            double TaskDataTableListCount = TaskDataTableList.Count();

            double CPTPercent = ((ActsWithoutPrecCount + ActsWithoutSuccCount) / TaskDataTableListCount) * 100;

            double CPTPercentRounded = Math.Round(CPTPercent, 2, MidpointRounding.AwayFromZero);

            string CPTPercentRoundedText = CPTPercentRounded.ToString() + "%";

            bool ispassed = true;

            if ((ActsWithoutSuccCount + ActsWithoutPrecCount) > 2)
            {
                ispassed = false;
            }


            return (TotalActsWithoutSuccORPred, CPTPercentRoundedText, ispassed);

        }






        // 13. CRITICAL PATH LENGTH INDEX (CPLI)

        public (string CPLIPercent, bool IsPassed) CPLICheck(DateTime? DataDate, DataTable TASKPREDDataTable,
            DataTable TaskDataTable, DataTable CalendarDataTable, XerFunctions xerfunc)
        {

            var CalendarDataTableList = (from c in CalendarDataTable.AsEnumerable()
                                         select new
                                         {
                                             day_hr_cnt = c.Field<string>("day_hr_cnt"),
                                             clndr_id = c.Field<string>("clndr_id")
                                         }).ToList();

            var TaskDataTableList = (from t in TaskDataTable.AsEnumerable()


                                     select new
                                     {

                                         target_end_date = xerfunc.Convertdatefromstring(t.Field<string>("target_end_date")),
                                         task_code = t.Field<string>("task_code"),
                                         task_name = t.Field<string>("task_name"),
                                         task_id = t.Field<string>("task_id"),
                                         total_float_hr_cnt = t.Field<string>("total_float_hr_cnt"),
                                         clndr_id = t.Field<string>("clndr_id"),
                                         wbs_id = t.Field<string>("wbs_id"),
                                         task_type = t.Field<string>("task_type")

                                     }).ToList().OrderBy(a => a.task_code);





            DateTime? MaxDate = TaskDataTableList.Select(a => a.target_end_date).Max();





            double RemainingDays = (MaxDate - DataDate).Value.Days;



            // Find TF :

            string LastAct = null;

            LastAct = TaskDataTableList.Where(a => a.task_type == "TT_FinMile")

              .Select(a => a.task_id).LastOrDefault();

            if (string.IsNullOrWhiteSpace(LastAct))
            {
                LastAct = TaskDataTableList.Where(a => a.target_end_date == MaxDate)

                 .Select(a => a.task_id).LastOrDefault();
            }






            string LastAct_wbs_id = TaskDataTableList.Where(a => a.task_id == LastAct)
                .Select(a => a.wbs_id).FirstOrDefault();

            var TASKPREDTableListofLast = (from t in TASKPREDDataTable.AsEnumerable()
                                           where t.Field<string>("task_id") == LastAct

                                           select new
                                           {

                                               pred_task_id = t.Field<string>("pred_task_id"),
                                               task_id = t.Field<string>("task_id")
                                           }).ToList();


            var PreListTF = (from t in TaskDataTableList
                             join p in TASKPREDTableListofLast
                             on t.task_id equals p.pred_task_id

                             select new
                             {
                                 total_float_hr_cnt = t.total_float_hr_cnt,
                                 clndr_id = t.clndr_id,
                                 task_id = t.task_id,
                                 wbs_id = t.wbs_id
                             }).ToList();



            var PreListTF_SameWBS = PreListTF.Where(a => a.wbs_id == LastAct_wbs_id).ToList();



            var PreTasksofLastTFDays = (from t in PreListTF_SameWBS
                                        join c in CalendarDataTableList
                                        on t.clndr_id equals c.clndr_id
                                        select new
                                        {
                                            wbs_id = t.wbs_id,
                                            task_id = t.task_id,
                                            TFinDays = xerfunc.ConvertFromHourtoDays(t.total_float_hr_cnt, c.day_hr_cnt)
                                        }).ToList();







            double TF_Final = PreTasksofLastTFDays.Select(a => a.TFinDays).DefaultIfEmpty().Max();



            double CPLI = Math.Round(((RemainingDays + TF_Final) / RemainingDays), 2, MidpointRounding.AwayFromZero);

            double CPLIabs = Math.Abs(CPLI);

            string CPLIPErcentText = CPLIabs.ToString();

            bool ispassed = true;

            if (CPLI < 0.95)
            {
                ispassed = false;
            }




            return (CPLIPErcentText, ispassed);



        }








        //14. BASELINE EXECUTION INDEX (BEI)


        public (List<(string, string)> CompletedActs, string BEIPercent, bool ispassed)
            BEICheck(DataTable TaskDataTable, XerFunctions xerfunc, DateTime? DataDate)

        {

            var TaskDataTableListCompleted = (from t in TaskDataTable.AsEnumerable()
                                              where t.Field<string>("status_code") == "TK_Complete"
                                              select new
                                              {
                                                  act_start_date = t.Field<string>("act_start_date"),
                                                  act_end_date = t.Field<string>("act_end_date"),
                                                  target_start_date = t.Field<string>("target_start_date"),
                                                  target_end_date = t.Field<string>("target_end_date"),
                                                  task_code = t.Field<string>("task_code"),
                                                  task_name = t.Field<string>("task_name")

                                              }).ToList().OrderBy(a => a.task_code);



            List<(string, string)> CompletedList = new List<(string, string)>();

            foreach (var item in TaskDataTableListCompleted)
            {
                CompletedList.Add((item.task_code, item.task_name));
            }


            var TaskDataTableList = (from t in TaskDataTable.AsEnumerable()

                                     select new
                                     {

                                         target_end_date = t.Field<string>("target_end_date"),
                                         task_code = t.Field<string>("task_code"),
                                         task_name = t.Field<string>("task_name")

                                     }).ToList().OrderBy(a => a.task_code);


            var TaskDataTableListConverted = from t in TaskDataTableList
                                             select new
                                             {

                                                 target_end_date = xerfunc.Convertdatefromstring(t.target_end_date),
                                                 task_code = t.task_code,
                                                 task_name = t.task_name
                                             };


            var TasksShouldBeComppleted = TaskDataTableListConverted.Where(a => a.target_end_date <= DataDate);


            double TaskDataTableListCompletedCount = TaskDataTableListCompleted.Count();


            double TasksShouldBeComppletedCount = TasksShouldBeComppleted.Count();

            double BEIPercent = (TaskDataTableListCompletedCount / TasksShouldBeComppletedCount);

            double BEIPercentRounded;

            if ((TasksShouldBeComppletedCount == 0) && (TaskDataTableListCompletedCount == 0))
            {
                BEIPercentRounded = 1;
            }


            else if (TasksShouldBeComppletedCount == 0)
            {
                BEIPercentRounded = 0;
            }
            else
            {


                BEIPercentRounded = Math.Round(BEIPercent, 0, MidpointRounding.AwayFromZero);
            }



            string BEIPercentRoundedText = BEIPercentRounded.ToString();

            bool ispassed = true;

            if (BEIPercentRounded < 0.95)
            {
                ispassed = false;
            }



            return (CompletedList, BEIPercentRoundedText, ispassed);



        }




        public void FinalDCMA14(string folderName, string XerFileContent, XerExtractor xerextractor, XerDCMACheckers xerdcma, XerFunctions xerfunc)

        {

            DataTable TaskDataTable = xerextractor.TASKTable(XerFileContent, xerextractor, xerfunc);

            string ProjectName = xerextractor.ProjectName(XerFileContent, xerextractor, xerfunc);

            DataTable CalendarDataTable = xerextractor.CalendarTable(XerFileContent, xerextractor, xerfunc);

            var TASKPREDDataTable = xerextractor.TASKPREDTable(XerFileContent, xerextractor, xerfunc);



            string datadate = xerextractor.GetDataDate(XerFileContent, xerextractor, xerfunc);

            var DataDate = xerfunc.Convertdatefromstring(datadate);


            string datetimenow = DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss");



            string pathString = System.IO.Path.Combine(folderName, " DCMA " + ProjectName + "  " + datetimenow);
            System.IO.Directory.CreateDirectory(pathString);



            string ImagePath = pathString + @"\" + "Mini Report.Jpeg";




            Dictionary<string, string> PercentDic = new Dictionary<string, string>();

            Dictionary<string, bool> IsPassedDic = new Dictionary<string, bool>();









            //1
            var MissingLogicCheckPercent = xerdcma.MissingLogicCheck(TaskDataTable, TASKPREDDataTable).LogicPercent;

            var MissingLogicCheckbool = xerdcma.MissingLogicCheck(TaskDataTable, TASKPREDDataTable).IsPassed;

            var MissedLogicList = xerdcma.MissingLogicCheck(TaskDataTable, TASKPREDDataTable).NonLogicActs;

           

            PercentDic.Add("Logic", MissingLogicCheckPercent);
            IsPassedDic.Add("Logic", MissingLogicCheckbool);

            //2
            var LeadsCheckPercent = xerdcma.LeadsCheck(TaskDataTable, TASKPREDDataTable, CalendarDataTable).LeadPercent;

            var LeadsCheckbool = xerdcma.LeadsCheck(TaskDataTable, TASKPREDDataTable, CalendarDataTable).IsPassed;

            var LeadsActs = xerdcma.LeadsCheck(TaskDataTable, TASKPREDDataTable, CalendarDataTable).PreActivitesWithLead;

          

            PercentDic.Add("Lead", LeadsCheckPercent);
            IsPassedDic.Add("Lead", LeadsCheckbool);


            //3
            var LagCheckPercent = xerdcma.LagCheck(TaskDataTable, TASKPREDDataTable, CalendarDataTable).LagPercent;

            var LagCheckbool = xerdcma.LagCheck(TaskDataTable, TASKPREDDataTable, CalendarDataTable).IsPassed;

            var LagActs = xerdcma.LagCheck(TaskDataTable, TASKPREDDataTable, CalendarDataTable).PreActivitesWithLag;



            PercentDic.Add("Lag", LagCheckPercent);
            IsPassedDic.Add("Lag", LagCheckbool);

            //4 
            var FSRelationshipCheckPercent = xerdcma.FSRelationshipCheck(TaskDataTable, TASKPREDDataTable).FSPercent;

            var FSRelationshipCheckbool = xerdcma.FSRelationshipCheck(TaskDataTable, TASKPREDDataTable).IsPassed;

            var NonFSActs = xerdcma.FSRelationshipCheck(TaskDataTable, TASKPREDDataTable).NonFSActivites;

          

            PercentDic.Add("FSRelations", FSRelationshipCheckPercent);
            IsPassedDic.Add("FSRelations", FSRelationshipCheckbool);


            //5 
            var HardConsCheckPercent = xerdcma.HardConsCheck(TaskDataTable).HardConsPercent;

            var HardConsCheckbool = xerdcma.HardConsCheck(TaskDataTable).IsPassed;

            var HardConsActs = xerdcma.HardConsCheck(TaskDataTable).HardConsActs;

         


            PercentDic.Add("HardCons", HardConsCheckPercent);
            IsPassedDic.Add("HardCons", HardConsCheckbool);



            //6 
            var HighFloatCheckPercent = xerdcma.HighFloatCheck(TaskDataTable, CalendarDataTable, xerfunc).HighFloatPercent;

            var HighFloatCheckbool = xerdcma.HighFloatCheck(TaskDataTable, CalendarDataTable, xerfunc).ispassed;

            var HighFloatActs = xerdcma.HighFloatCheck(TaskDataTable, CalendarDataTable, xerfunc).HighFloatActs;

           


            PercentDic.Add("HighFloat", HighFloatCheckPercent);
            IsPassedDic.Add("HighFloat", HighFloatCheckbool);




            //7
            var NegativeFloatCheckPercent = xerdcma.NegativeFloatCheck(TaskDataTable, CalendarDataTable, xerfunc).NegativeFloatPercent;

            var NegativeFloatCheckbool = xerdcma.NegativeFloatCheck(TaskDataTable, CalendarDataTable, xerfunc).ispassed;

            var NegativeFloatActs = xerdcma.NegativeFloatCheck(TaskDataTable, CalendarDataTable, xerfunc).NegativeFloatActs;

           

            PercentDic.Add("NegativeFloat", NegativeFloatCheckPercent);
            IsPassedDic.Add("NegativeFloat", NegativeFloatCheckbool);

            //8
            var HighDurationCheckPercent = xerdcma.HighDurationCheck(TaskDataTable, CalendarDataTable, xerfunc).HighDurationPercent;

            var HighDurationCheckbool = xerdcma.HighDurationCheck(TaskDataTable, CalendarDataTable, xerfunc).ispassed;

            var HighDurationActs = xerdcma.HighDurationCheck(TaskDataTable, CalendarDataTable, xerfunc).HighDurationActs;

       

            PercentDic.Add("HighDur", HighDurationCheckPercent);
            IsPassedDic.Add("HighDur", HighDurationCheckbool);

            //9
            var InvalidDatesCheckPercent = xerdcma.InvalidDatesCheck(TaskDataTable, DataDate, xerfunc).InvalidDatesPercent;

            var InvalidDatesCheckbool = xerdcma.InvalidDatesCheck(TaskDataTable, DataDate, xerfunc).ispassed;

            var InvalidDatesActs = xerdcma.InvalidDatesCheck(TaskDataTable, DataDate, xerfunc).InvalidDatesActs;

        


            PercentDic.Add("InvaliDates", InvalidDatesCheckPercent);
            IsPassedDic.Add("InvaliDates", InvalidDatesCheckbool);




            //10 

            var MissedResourcesCheckPercent = xerdcma.MissedResourcesCheck(TaskDataTable).MissedResourcesPercent;

            var MissedResourcesCheckbool = xerdcma.MissedResourcesCheck(TaskDataTable).ispassed;

            var MissedResourcesActs = xerdcma.MissedResourcesCheck(TaskDataTable).ActswithMissedResources;

      



            PercentDic.Add("Resources", MissedResourcesCheckPercent);
            IsPassedDic.Add("Resources", MissedResourcesCheckbool);

            //11 

            var MissedTasksCheckPercent = xerdcma.MissedTasksCheck(TaskDataTable, DataDate, xerfunc).MissedTaskPercent;

            var MissedTasksCheckbool = xerdcma.MissedTasksCheck(TaskDataTable, DataDate, xerfunc).ispassed;

            var MissedActs = xerdcma.MissedTasksCheck(TaskDataTable, DataDate, xerfunc).MissedList;

         


            PercentDic.Add("MissedTasks", MissedTasksCheckPercent);
            IsPassedDic.Add("MissedTasks", MissedTasksCheckbool);



            //12 


            var CPTCheckPercent = xerdcma.CPTCheck(TaskDataTable, TASKPREDDataTable).ActsPercent;

            var CPTCheckbool = xerdcma.CPTCheck(TaskDataTable, TASKPREDDataTable).ispassed;

            var ActsWithSuccorPrec = xerdcma.CPTCheck(TaskDataTable, TASKPREDDataTable).ActsWithoutList;

           

            PercentDic.Add("CPT", CPTCheckPercent);
            IsPassedDic.Add("CPT", CPTCheckbool);


            //13 

            var CPLICheckPercent = xerdcma.CPLICheck(DataDate, TASKPREDDataTable, TaskDataTable, CalendarDataTable, xerfunc).CPLIPercent;

            var CPLICheckbool = xerdcma.CPLICheck(DataDate, TASKPREDDataTable, TaskDataTable, CalendarDataTable, xerfunc).IsPassed;

        


            PercentDic.Add("CPLI", CPLICheckPercent);
            IsPassedDic.Add("CPLI", CPLICheckbool);



            //14
            var BEIPercent = xerdcma.BEICheck(TaskDataTable, xerfunc, DataDate).BEIPercent;

            var BEIbool = xerdcma.BEICheck(TaskDataTable, xerfunc, DataDate).ispassed;

          


            PercentDic.Add("BEI", BEIPercent);
            IsPassedDic.Add("BEI", BEIbool);


            // Reporting Image + excelsheet 

            xerfunc.MiniImageDCMA(ImagePath, ProjectName, IsPassedDic, PercentDic);




            string excelfilepath = pathString + @"\" + "Report.xlsx";
            var fileexcel = new FileInfo(excelfilepath);
            if (fileexcel.Exists)
                fileexcel.Delete();

            var pck = new ExcelPackage(fileexcel);
            var workbook = pck.Workbook;



            xerfunc.ExcelSheetDCMAFunc(workbook, "Missing Logic Activites", MissedLogicList);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Acitvites with Lead", LeadsActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Acitvites with Lag", LagActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Non FS Relationship Activites", NonFSActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Activites with Hard Constraint", HardConsActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Acitvites with High Float", HighFloatActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Acitvites with Negative Float", NegativeFloatActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Acitvites with High Duration", HighDurationActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Acitvites with Invalid Dates", InvalidDatesActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Acitvites with Missing Resources", MissedResourcesActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Missed Target Dates Acitvites", MissedActs);
            xerfunc.ExcelSheetDCMAFunc(workbook, "Acitvites without Successer or Prec", ActsWithSuccorPrec);





            pck.Save();

        }









    }



}
