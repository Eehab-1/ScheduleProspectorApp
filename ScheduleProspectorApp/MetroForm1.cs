using OfficeOpenXml;
using Syncfusion.WinForms.Controls;
using ScheduleProspectorApp.MyClasses.XER;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace ScheduleProspectorApp
{
    public partial class MetroForm1 : SfForm
    {
        public MetroForm1()
        {
            InitializeComponent();

            this.Style.TitleBar.Height = 50;
            this.Style.TitleBar.BackColor = Color.Blue;
            this.Style.TitleBar.ForeColor = Color.White;
            this.Style.TitleBar.IconBackColor = Color.FromArgb(15, 161, 212);
            this.BackColor = Color.SkyBlue;
            //this.Style.TitleBar.ForeColor = ColorTranslator.FromHtml("#343434");
            this.Style.TitleBar.BackColor = Color.FromArgb(15, 161, 212);

            this.Style.TitleBar.CloseButtonForeColor = Color.White;
            this.Style.TitleBar.MaximizeButtonForeColor = Color.White;
            this.Style.TitleBar.MinimizeButtonForeColor = Color.White;
            this.Style.TitleBar.HelpButtonForeColor = Color.DarkGray;
            this.Style.TitleBar.IconHorizontalAlignment = HorizontalAlignment.Left;
            this.Style.TitleBar.Font = this.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Style.TitleBar.TextHorizontalAlignment = HorizontalAlignment.Left;
            this.Style.TitleBar.TextVerticalAlignment = System.Windows.Forms.VisualStyles.VerticalAlignment.Center;
           // this.CaptionBarColor = Color.Pink;
           
        }

        private void MetroForm1_Load(object sender, EventArgs e)
        {

        }

        private void tabControlAdv1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabControlAdv1_Enter(object sender, EventArgs e)
        {
          
        }

        string filecontent;
        private async void ImportBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
               openFileDialog1.Filter = "File File (*.xer)|*.xer";
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;
            string path = openFileDialog1.FileName;


            await Task.Run(() =>
            {
                filecontent = File.ReadAllText(path, Encoding.UTF8);
                ProjectNamelabel.Text = path;
            });
        }

        private void ExportBtn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {

                Backuplabel.Text = folderBrowserDialog1.SelectedPath;
               

               

            }
        }

        private async void sfButton1_Click(object sender, EventArgs e)
        {
           

            if (string.IsNullOrWhiteSpace(filecontent))
            {
                MessageBox.Show(" You MUST import a file  !", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (Backuplabel.Text== "No output chosen")
            {
                MessageBox.Show(" You MUST choose Backup folder !", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }



            string folderName = Backuplabel.Text;

            string XerFileContent = filecontent;

            XerExtractor xerextractor = new XerExtractor();

            XerDCMACheckers xerdcma = new XerDCMACheckers();

            XerFunctions xerfunc = new XerFunctions();


            if (radioButton1.Checked == true)
            {
                await Task.Run(() =>
                {
                    try
                    {
                        xerfunc.GlanceAssessment(folderName, XerFileContent, xerextractor, xerdcma, xerfunc);
                    }

                    catch
                    {

                    }
                    
                });
            }

            else
            {


                await Task.Run(() =>
                {
                    try
                    {
                        xerdcma.FinalDCMA14(folderName, XerFileContent, xerextractor, xerdcma, xerfunc);
                    }


                    catch
                    {

                    }



                });



            }



            await Task.Delay(500);

            await Task.Run(() =>
            {
                try
                {
                    Process.Start(Backuplabel.Text);
                }

                catch { }

            });



        }

        private async void sfButton2_Click(object sender, EventArgs e)
        {


            if (string.IsNullOrWhiteSpace(filecontent))
            {
                MessageBox.Show(" You MUST import a file  !", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (Backuplabel.Text == "No output chosen")
            {
                MessageBox.Show(" You MUST choose Backup folder !", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }





            string XerFileContent = filecontent;

            //  string XerFileContent = string.Join("", GlobalXMLList);

            XerExtractor xerextractor = new XerExtractor();

            XerDCMACheckers xerdcma = new XerDCMACheckers();

            XerFunctions xerfunc = new XerFunctions();


            await Task.Run(() =>
            {


                try
                {




                    DataTable TaskDataTable = xerextractor.TASKTable(XerFileContent, xerextractor, xerfunc);

                    DataTable CalendarDataTable = xerextractor.CalendarTable(XerFileContent, xerextractor, xerfunc);

                    var TASKPREDDataTable = xerextractor.TASKPREDTable(XerFileContent, xerextractor, xerfunc);


                    string ProjectName = xerextractor.ProjectName(XerFileContent, xerextractor, xerfunc);


                    var TaskDataTableList = (from t in TaskDataTable.AsEnumerable()

                                             select new
                                             {
                                                 task_id = t.Field<string>("task_id"),
                                                 task_code = t.Field<string>("task_code"),
                                                 task_name = t.Field<string>("task_name")

                                             }).ToList().OrderBy(a => a.task_code);




                    foreach (DataRow row in TASKPREDDataTable.Rows)
                    {
                        if (row["task_id"].ToString() != "task_id")
                        {

                            string task_code = TaskDataTableList.Where(a => a.task_id == row["task_id"].ToString())
                                .Select(a => a.task_code).FirstOrDefault();

                            row["task_id"] = task_code;
                        }

                        if (row["pred_task_id"].ToString() != "task_id")
                        {

                            string task_code = TaskDataTableList.Where(a => a.task_id == row["pred_task_id"].ToString())
                                .Select(a => a.task_code).FirstOrDefault();

                            row["pred_task_id"] = task_code;
                        }



                    }








                    string datetimenow = DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss");

                    string folderName = Backuplabel.Text;

                    string pathString = System.IO.Path.Combine(folderName, ProjectName + " Extracted "+ "  " + datetimenow);
                    System.IO.Directory.CreateDirectory(pathString);

                    string excelfilepath = pathString + @"\" + "Extracted_Data.xlsx";


                    using (ExcelPackage pck = new ExcelPackage(excelfilepath))
                    {
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Activites");
                        ws.Cells["A1"].LoadFromDataTable(TaskDataTable, true);

                        ExcelWorksheet ws2 = pck.Workbook.Worksheets.Add("Relationships");
                        ws2.Cells["A1"].LoadFromDataTable(TASKPREDDataTable, true);

                        ExcelWorksheet ws3 = pck.Workbook.Worksheets.Add("Calendars");
                        ws3.Cells["A1"].LoadFromDataTable(CalendarDataTable, true);



                        pck.Save();
                    }


                    

                }


                catch { }


            });

            await Task.Delay(500);

            await Task.Run(() =>
            {
                try
                {
                    Process.Start(Backuplabel.Text);
                }
               
                catch { } 
                
            });


        }

        private async  void sfButton3_Click(object sender, EventArgs e)
        {
            await Task.Run(() =>
            {
                Process.Start("http://scheduleprospector.com");
            });
           
        }

        private async void sfButton4_Click(object sender, EventArgs e)
        {
            await Task.Run(() =>
            {
                Process.Start("ms-windows-store:updates");
            });
           


        }

        private void tabPageAdv2_Click(object sender, EventArgs e)
        {

        }
    }
}
