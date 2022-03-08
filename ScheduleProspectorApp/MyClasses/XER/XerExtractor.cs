using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ScheduleProspectorApp.MyClasses.XER
{

    public class XerExtractor
    {


        public List<string> GetTables(string XerFileContent)
        {


            XerFileContent += "\n" + "%T";

            string pattern = "%T";

            Regex regex = new Regex(pattern);

            var matches = regex.Matches(XerFileContent);


            List<int> Tindexes = new List<int>();

            foreach (Match match in matches)
            {


                Tindexes.Add(match.Index);
            }



            List<string> TablesList = new List<string>();

            for (int i = 0; i < Tindexes.Count; i++)
            {
                if (Tindexes[i] != Tindexes[Tindexes.Count - 1])
                {


                    string table = XerFileContent.Substring(Tindexes[i], Tindexes[i + 1] - Tindexes[i]);



                    TablesList.Add(table);
                }

            }



            return TablesList;


        }


        public string GetDataDate(string XerFileContent, XerExtractor xerextractor, XerFunctions xerfunc)

        {


            var Tables = xerextractor.GetTables(XerFileContent);

            string ProjectTable = Tables.Where(a => a.Contains("last_recalc_date")).FirstOrDefault();

            List<string> ProjectTableList = xerfunc.TableStringToList(ProjectTable);

            ProjectTableList = ProjectTableList.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();


            ProjectTableList.RemoveAt(0);


            List<string> last_recalc_date_list = new List<string>();

            List<string> plan_start_date_List = new List<string>();

            List<int> last_recalc_date_Index_List = new List<int>();

            List<int> plan_start_date_Index_List = new List<int>();



            foreach (var line in ProjectTableList)
            {




                var delimiters = new char[] { '\t' };


                string[] segments = line.Split(delimiters);

                var last_recalc_date_index = segments.ToList().IndexOf("last_recalc_date");

                var plan_start_date_index = segments.ToList().IndexOf("plan_start_date");

                last_recalc_date_Index_List.Add(last_recalc_date_index);

                plan_start_date_Index_List.Add(plan_start_date_index);

                // richTextBox6.AppendText(last_recalc_date_Index_List[0].ToString() +Environment.NewLine);

                for (int i = 0; i < segments.Length; i++)
                {




                    if (i == last_recalc_date_Index_List[0])
                    {
                        //   richTextBox2.AppendText(segments[i] + Environment.NewLine);
                        if (!string.IsNullOrWhiteSpace(segments[i]))
                        {
                            last_recalc_date_list.Add(segments[i]);
                        }

                    }

                    if (i == plan_start_date_Index_List[0])
                    {
                        if (!string.IsNullOrWhiteSpace(segments[i]))
                        {
                            plan_start_date_List.Add(segments[i]);
                        }

                    }

                }

            }


            var plan_start_date_Converted = xerfunc.Convertdatefromstring(plan_start_date_List[plan_start_date_List.Count - 1]);

            var last_recalc_date_Converted = xerfunc.Convertdatefromstring(last_recalc_date_list[last_recalc_date_list.Count - 1]);

            if (last_recalc_date_Converted > plan_start_date_Converted)
            {


                return last_recalc_date_list[last_recalc_date_list.Count - 1];
            }

            else
            {
                return plan_start_date_List[plan_start_date_List.Count - 1];
            }





        }


        public DataTable TASKPREDTable(string XerFileContent, XerExtractor xerextractor, XerFunctions xerfunctions)

        {





            var Tables = xerextractor.GetTables(XerFileContent);

            string RelationsTable = Tables.Where(a => a.Contains("task_pred_id")).FirstOrDefault();

            List<string> RelationsTableList = xerfunctions.TableStringToList(RelationsTable);

            RelationsTableList = RelationsTableList.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();


            RelationsTableList.RemoveAt(0);


            List<string> task_id_List = new List<string>();

            List<string> pred_task_id_List = new List<string>();

            List<string> pred_type_List = new List<string>();

            List<string> lag_hr_cnt_List = new List<string>();


            List<int> task_id_index_List = new List<int>();

            List<int> pred_task_id_index_List = new List<int>();

            List<int> pred_type_index_List = new List<int>();

            List<int> lag_hr_cnt_index_List = new List<int>();





            foreach (var line in RelationsTableList)
            {




                var delimiters = new char[] { '\t' };


                string[] segments = line.Split(delimiters);


                int task_id_index = segments.ToList().IndexOf("task_id");

                task_id_index_List.Add(task_id_index);


                int pred_task_id_index = segments.ToList().IndexOf("pred_task_id");

                pred_task_id_index_List.Add(pred_task_id_index);


                int pred_type_index = segments.ToList().IndexOf("pred_type");

                pred_type_index_List.Add(pred_type_index);


                int lag_hr_cnt_index = segments.ToList().IndexOf("lag_hr_cnt");

                lag_hr_cnt_index_List.Add(lag_hr_cnt_index);



                for (int i = 0; i < segments.Length; i++)
                {




                    if (i == task_id_index_List[0])
                    {


                        task_id_List.Add(segments[i]);
                    }


                    if (i == pred_task_id_index_List[0])
                    {


                        pred_task_id_List.Add(segments[i]);
                    }

                    if (i == pred_type_index_List[0])
                    {


                        pred_type_List.Add(segments[i]);
                    }

                    if (i == lag_hr_cnt_index_List[0])
                    {


                        lag_hr_cnt_List.Add(segments[i]);
                    }


                }



            }


            task_id_List.RemoveAt(0);
            pred_task_id_List.RemoveAt(0);
            pred_type_List.RemoveAt(0);
            lag_hr_cnt_List.RemoveAt(0);






            DataTable datatable = new DataTable();

            datatable.Columns.Add("task_id", typeof(string));
            datatable.Columns.Add("pred_task_id", typeof(string));
            datatable.Columns.Add("pred_type", typeof(string));
            datatable.Columns.Add("lag_hr_cnt", typeof(string));

            for (int i = 0; i < task_id_List.Count; i++)
            {


                datatable.Rows.Add(task_id_List[i], pred_task_id_List[i], pred_type_List[i], lag_hr_cnt_List[i]);

            }






            return datatable;


        }


        public DataTable TASKTable(string XerFileContent, XerExtractor xerextractor, XerFunctions xerfunctions)
        {






            var Tables = xerextractor.GetTables(XerFileContent);

            string TaskTable = Tables.Where(a => a.Contains("total_float_hr_cnt")).FirstOrDefault();

            List<string> TaskTableList = xerfunctions.TableStringToList(TaskTable);

            TaskTableList = TaskTableList.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();


            TaskTableList.RemoveAt(0);


            List<string> task_id_List = new List<string>();

            List<string> clndr_id_List = new List<string>();

            List<string> phys_complete_pct_List = new List<string>();

            List<string> task_type_List = new List<string>();

            List<string> status_code_List = new List<string>();

            List<string> task_code_List = new List<string>();

            List<string> task_name_List = new List<string>();

            List<string> rsrc_id_List = new List<string>();

            List<string> total_float_hr_cnt_List = new List<string>();

            List<string> remain_drtn_hr_cnt_List = new List<string>();

            List<string> target_drtn_hr_cnt_List = new List<string>();

            List<string> cstr_type_List = new List<string>();

            List<string> act_start_date_List = new List<string>();

            List<string> act_end_date_List = new List<string>();

            List<string> target_start_date_List = new List<string>();

            List<string> target_end_date_List = new List<string>();

            List<string> driving_path_flag_List = new List<string>();

            List<string> wbs_id_List = new List<string>();


            List<int> task_id_index_List = new List<int>();

            List<int> clndr_id_index_List = new List<int>();

            List<int> phys_complete_pct_index_List = new List<int>();

            List<int> task_type_index_List = new List<int>();

            List<int> status_code_index_List = new List<int>();

            List<int> task_code_index_List = new List<int>();

            List<int> task_name_index_List = new List<int>();

            List<int> rsrc_id_index_List = new List<int>();

            List<int> total_float_hr_cnt_index_List = new List<int>();

            List<int> remain_drtn_hr_cnt_index_List = new List<int>();

            List<int> target_drtn_hr_cnt_index_List = new List<int>();

            List<int> cstr_type_index_List = new List<int>();

            List<int> act_start_date_index_List = new List<int>();

            List<int> act_end_date_index_List = new List<int>();

            List<int> target_start_date_index_List = new List<int>();

            List<int> target_end_date_index_List = new List<int>();

            List<int> driving_path_flag_index_List = new List<int>();

            List<int> wbs_id_index_List = new List<int>();






            foreach (var line in TaskTableList)
            {




                var delimiters = new char[] { '\t' };


                string[] segments = line.Split(delimiters);

                int task_id_index = segments.ToList().IndexOf("task_id");

                task_id_index_List.Add(task_id_index);


                int clndr_id_index = segments.ToList().IndexOf("clndr_id");

                clndr_id_index_List.Add(clndr_id_index);


                int phys_complete_pct_index = segments.ToList().IndexOf("phys_complete_pct");

                phys_complete_pct_index_List.Add(phys_complete_pct_index);


                int task_type_index = segments.ToList().IndexOf("task_type");

                task_type_index_List.Add(task_type_index);


                int status_code_index = segments.ToList().IndexOf("status_code");

                status_code_index_List.Add(status_code_index);


                int task_code_index = segments.ToList().IndexOf("task_code");

                task_code_index_List.Add(task_code_index);


                int task_name_index = segments.ToList().IndexOf("task_name");

                task_name_index_List.Add(task_name_index);


                int rsrc_id_index = segments.ToList().IndexOf("rsrc_id");

                rsrc_id_index_List.Add(rsrc_id_index);


                int total_float_hr_cnt_index = segments.ToList().IndexOf("total_float_hr_cnt");

                total_float_hr_cnt_index_List.Add(total_float_hr_cnt_index);


                int remain_drtn_hr_cnt_index = segments.ToList().IndexOf("remain_drtn_hr_cnt");

                remain_drtn_hr_cnt_index_List.Add(remain_drtn_hr_cnt_index);


                int target_drtn_hr_cnt_index = segments.ToList().IndexOf("target_drtn_hr_cnt");

                target_drtn_hr_cnt_index_List.Add(target_drtn_hr_cnt_index);

                int cstr_type_index = segments.ToList().IndexOf("cstr_type");

                cstr_type_index_List.Add(cstr_type_index);


                int act_start_date_index = segments.ToList().IndexOf("act_start_date");

                act_start_date_index_List.Add(act_start_date_index);

                int act_end_date_index = segments.ToList().IndexOf("act_end_date");

                act_end_date_index_List.Add(act_end_date_index);


                int target_start_date_index = segments.ToList().IndexOf("target_start_date");

                target_start_date_index_List.Add(target_start_date_index);



                int target_end_date_index = segments.ToList().IndexOf("target_end_date");

                target_end_date_index_List.Add(target_end_date_index);


                int driving_path_flag_index = segments.ToList().IndexOf("driving_path_flag");

                driving_path_flag_index_List.Add(driving_path_flag_index);



                int wbs_id_index = segments.ToList().IndexOf("wbs_id");

                wbs_id_index_List.Add(wbs_id_index);


                for (int i = 0; i < segments.Length; i++)
                {




                    if (i == task_id_index_List[0])
                    {


                        task_id_List.Add(segments[i]);
                    }

                    if (i == wbs_id_index_List[0])
                    {
                        wbs_id_List.Add(segments[i]);
                    }


                    if (i == clndr_id_index_List[0])
                    {


                        clndr_id_List.Add(segments[i]);
                    }

                    if (i == phys_complete_pct_index_List[0])
                    {


                        phys_complete_pct_List.Add(segments[i]);
                    }

                    if (i == task_type_index_List[0])
                    {


                        task_type_List.Add(segments[i]);
                    }



                    if (i == status_code_index_List[0])
                    {


                        status_code_List.Add(segments[i]);
                    }

                    if (i == task_code_index_List[0])
                    {


                        task_code_List.Add(segments[i]);
                    }

                    if (i == task_name_index_List[0])
                    {


                        task_name_List.Add(segments[i]);
                    }

                    if (i == rsrc_id_index_List[0])
                    {


                        rsrc_id_List.Add(segments[i]);
                    }

                    if (i == total_float_hr_cnt_index_List[0])
                    {


                        total_float_hr_cnt_List.Add(segments[i]);
                    }

                    if (i == remain_drtn_hr_cnt_index_List[0])
                    {


                        remain_drtn_hr_cnt_List.Add(segments[i]);
                    }

                    if (i == target_drtn_hr_cnt_index_List[0])
                    {


                        target_drtn_hr_cnt_List.Add(segments[i]);
                    }

                    if (i == cstr_type_index_List[0])
                    {


                        cstr_type_List.Add(segments[i]);
                    }

                    if (i == driving_path_flag_index_List[0])
                    {
                        driving_path_flag_List.Add(segments[i]);
                    }


                    if (i == target_end_date_index_List[0])
                    {
                        target_end_date_List.Add(segments[i]);
                    }

                    if (i == target_start_date_index_List[0])
                    {
                        target_start_date_List.Add(segments[i]);
                    }

                    if (i == act_end_date_index_List[0])
                    {
                        act_end_date_List.Add(segments[i]);
                    }

                    if (i == act_start_date_index_List[0])
                    {
                        act_start_date_List.Add(segments[i]);
                    }

                }



            }




            task_id_List.RemoveAt(0);
            clndr_id_List.RemoveAt(0);
            phys_complete_pct_List.RemoveAt(0);
            task_type_List.RemoveAt(0);
            status_code_List.RemoveAt(0);
            task_code_List.RemoveAt(0);
            task_name_List.RemoveAt(0);
            rsrc_id_List.RemoveAt(0);
            total_float_hr_cnt_List.RemoveAt(0);
            remain_drtn_hr_cnt_List.RemoveAt(0);
            target_drtn_hr_cnt_List.RemoveAt(0);
            cstr_type_List.RemoveAt(0);
            act_start_date_List.RemoveAt(0);
            act_end_date_List.RemoveAt(0);
            target_start_date_List.RemoveAt(0);
            target_end_date_List.RemoveAt(0);
            driving_path_flag_List.RemoveAt(0);
            wbs_id_List.RemoveAt(0);

            DataTable datatable = new DataTable();

            datatable.Columns.Add("task_id", typeof(string));
            datatable.Columns.Add("clndr_id", typeof(string));
            datatable.Columns.Add("phys_complete_pct", typeof(string));
            datatable.Columns.Add("task_type", typeof(string));

            datatable.Columns.Add("status_code", typeof(string));
            datatable.Columns.Add("task_code", typeof(string));
            datatable.Columns.Add("task_name", typeof(string));
            datatable.Columns.Add("rsrc_id", typeof(string));

            datatable.Columns.Add("total_float_hr_cnt", typeof(string));
            datatable.Columns.Add("remain_drtn_hr_cnt", typeof(string));
            datatable.Columns.Add("target_drtn_hr_cnt", typeof(string));
            datatable.Columns.Add("cstr_type", typeof(string));

            datatable.Columns.Add("act_start_date", typeof(string));
            datatable.Columns.Add("act_end_date", typeof(string));
            datatable.Columns.Add("target_start_date", typeof(string));
            datatable.Columns.Add("target_end_date", typeof(string));

            datatable.Columns.Add("driving_path_flag", typeof(string));

            datatable.Columns.Add("wbs_id", typeof(string));

            for (int i = 0; i < task_id_List.Count; i++)
            {


                datatable.Rows.Add(task_id_List[i], clndr_id_List[i], phys_complete_pct_List[i], task_type_List[i]
                    , status_code_List[i], task_code_List[i], task_name_List[i], rsrc_id_List[i], total_float_hr_cnt_List[i]
                   , remain_drtn_hr_cnt_List[i], target_drtn_hr_cnt_List[i], cstr_type_List[i],
                   act_start_date_List[i], act_end_date_List[i], target_start_date_List[i], target_end_date_List[i],
                   driving_path_flag_List[i], wbs_id_List[i]);

            }







            return datatable;



        }


        public string ProjectName(string XerFileContent, XerExtractor xerextractor, XerFunctions xerfunctions)
        {

            //wbs_name
            //PROJWBS



            var Tables = xerextractor.GetTables(XerFileContent);

            string ProjectWBSTable = Tables.Where(a => a.Contains("wbs_name")).FirstOrDefault();

            List<string> ProjectWBSTableList = xerfunctions.TableStringToList(ProjectWBSTable);

            ProjectWBSTableList = ProjectWBSTableList.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();


            ProjectWBSTableList.RemoveAt(0);


            List<string> wbs_name_List = new List<string>();

            List<int> wbs_name_index_List = new List<int>();


            foreach (var line in ProjectWBSTableList)
            {




                var delimiters = new char[] { '\t' };


                string[] segments = line.Split(delimiters);


                int wbs_name_index = segments.ToList().IndexOf("wbs_name");

                wbs_name_index_List.Add(wbs_name_index);




                for (int i = 0; i < segments.Length; i++)
                {




                    if (i == wbs_name_index_List[0])
                    {


                        wbs_name_List.Add(segments[i]);
                    }





                }

            }


            wbs_name_List.RemoveAt(0);


            return wbs_name_List[0];


        }



        public DataTable CalendarTable(string XerFileContent, XerExtractor xerextractor, XerFunctions xerfunctions)
        {




            var Tables = xerextractor.GetTables(XerFileContent);

            string CALENDARTable = Tables.Where(a => a.Contains("day_hr_cnt")).FirstOrDefault();

            List<string> CALENDARTableList = xerfunctions.TableStringToList(CALENDARTable);

            CALENDARTableList = CALENDARTableList.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();


            CALENDARTableList.RemoveAt(0);


            List<string> clndr_id_List = new List<string>();

            List<string> clndr_name_List = new List<string>();

            List<string> day_hr_cnt_List = new List<string>();

            List<string> default_flag_List = new List<string>();


            List<int> clndr_id_index_List = new List<int>();

            List<int> clndr_name_index_List = new List<int>();

            List<int> day_hr_cnt_index_List = new List<int>();

            List<int> default_flag_index_List = new List<int>();




            foreach (var line in CALENDARTableList)
            {




                var delimiters = new char[] { '\t' };


                string[] segments = line.Split(delimiters);


                int clndr_id_index = segments.ToList().IndexOf("clndr_id");

                clndr_id_index_List.Add(clndr_id_index);

                int clndr_name_index = segments.ToList().IndexOf("clndr_name");

                clndr_name_index_List.Add(clndr_name_index);

                int day_hr_cnt_index = segments.ToList().IndexOf("day_hr_cnt");

                day_hr_cnt_index_List.Add(day_hr_cnt_index);

                int default_flag_index = segments.ToList().IndexOf("default_flag");

                default_flag_index_List.Add(default_flag_index);


                for (int i = 0; i < segments.Length; i++)
                {




                    if (i == clndr_id_index_List[0])
                    {


                        clndr_id_List.Add(segments[i]);
                    }

                    if (i == default_flag_index_List[0])
                    {
                        default_flag_List.Add(segments[i]);
                    }

                    if (i == clndr_name_index_List[0])
                    {


                        clndr_name_List.Add(segments[i]);
                    }

                    if (i == day_hr_cnt_index_List[0])
                    {


                        day_hr_cnt_List.Add(segments[i]);
                    }



                }

            }


            clndr_id_List.RemoveAt(0);
            clndr_name_List.RemoveAt(0);
            day_hr_cnt_List.RemoveAt(0);
            default_flag_List.RemoveAt(0);
            //  clndr_data_List.RemoveAt(0);

            DataTable datatable = new DataTable();

            datatable.Columns.Add("clndr_id", typeof(string));
            datatable.Columns.Add("clndr_name", typeof(string));
            datatable.Columns.Add("day_hr_cnt", typeof(string));
            datatable.Columns.Add("default_flag", typeof(string));
            // datatable.Columns.Add("clndr_data", typeof(string));




            for (int i = 0; i < clndr_id_List.Count; i++)
            {


                datatable.Rows.Add(clndr_id_List[i], clndr_name_List[i], day_hr_cnt_List[i]
                    , default_flag_List[i]);

            }




            return datatable;



        }



    }




}
