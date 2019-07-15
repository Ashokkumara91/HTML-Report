using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Collections;
using System.Net.Mail;
using System.Net.Mime;
using System.Threading;
using System.Data.SqlClient;
using System.Data;
using NReco.ImageGenerator;

namespace HTMLReport
{
    public class Report
    {

        public static string ExecutionStartTime = System.DateTime.Now.ToString("hh:mm:ss.fff");
        public static string ExecutionStartTime2 = System.DateTime.Now.ToString("hhmmssfff");
        public static string ExecutionStartTime1 = System.DateTime.Now.ToString("HH:mm:ss.fff");
        public static string ExecutionEndTime, ExecutionEndTime1, Reportlocation;
        public static string folderpath; 
        public static string mplocation;
        //public static string mplocation1 = folderpath + "\\OverallStatus.html";
        public static string maillocation = folderpath + "\\Mail.html";
        public static string Emaillocation = folderpath + "\\EMail.jpeg";
        public static string loglocation ;
        public static StringBuilder Rows, Iteration, Log, MailBody;
        public string ActualResult, Message, Loginfo, Feildset, lineNo, EXType;
        public static string link, totalcalc, Text, OverallEndTime, OverallStartTime, filelocation3, ProjName, ApplicationName, LURL, IterationStartTime, IterationEndTime, Iterationduration, TestCaseID, Scenario, UserName, OSName, BrowserName, ExpectedResulted, ExpectedStatus, TestCaseDescription;
        public static int count, stepsNE1, count1, g1, RowCount, totalsteps = 0, failedsteps = 0, passedsteps = 0, starttimecount;
        public static TimeSpan duration, Iterationdue;
        private int passed = 0, failed = 0;
        public int Percentage;

        public List<string> StartTime = new List<string>();
        public List<string> EndTime = new List<string>();
        public List<string> MethodName = new List<string>();
        public List<string> Status = new List<string>();
        public List<string> Table = new List<string>();
        public List<string> Iterationtable = new List<string>();

        public void ProjectDetails(string projectname, string browser, string URL, string Application)
        {
            //For getting Project Name
            ProjName = projectname;
            //For getting login username
            UserName = Environment.UserName;
            //For getting OS name
            OSName = new Microsoft.VisualBasic.Devices.ComputerInfo().OSFullName;
            string browser1 = browser.ToLower();
            if (browser1.Contains("chrome")) { BrowserName = "Chrome"; }
            if (browser1.Contains("firefox")) { BrowserName = "firefox"; }
            if (browser1.Contains("ie")) { BrowserName = "Internet Explorer"; }
            
            LURL = URL;
            ApplicationName = Application;

            if (!Directory.Exists(folderpath))
            {
                Directory.CreateDirectory(folderpath);
            }
        }
        public void OverAllStartTime()
        {
            OverallStartTime = System.DateTime.Now.ToString("hh:mm:ss.fff");
        }
        public void IterationStartFunction(int IterationCount, string TestcaseID, string scenario, int totalrows, string expectedstatus, string expectedresulted, string testcasedescription, string ProjectLocation, string ReportLocation)
        {
            g1 = IterationCount;
            Scenario = scenario;
            TestCaseID = TestcaseID;
            RowCount = totalrows;
            ExpectedStatus = expectedstatus;
            TestCaseDescription = testcasedescription;
            ExpectedResulted = expectedresulted;

            totalcalc = ProjectLocation +"\\UnitTest1.cs";
            Text = File.ReadAllText(totalcalc);
            count1 = Regex.Matches(Text, "//HTMLreport.StepStartFunction").Count;
            count = Regex.Matches(Text, "HTMLreport.StepStartFunction").Count;
            totalsteps = count - count1;

            folderpath = ReportLocation + "\\Report" + ExecutionStartTime2;

        }
        public void StepStartFunction(string name)
        {
            StartTime.Add(System.DateTime.Now.ToString("hh:mm:ss.fff"));
            MethodName.Add(name);
        }
        public void StepEndFunctionPassed(string sclink)
        {
            EndTime.Add(System.DateTime.Now.ToString("hh:mm:ss.fff"));
            Status.Add("Pass");
            link = sclink;
        }
        public void StepEndFunctionFailed(string sclink, string lineno, string Msg, string type)
        {
            EndTime.Add(System.DateTime.Now.ToString("hh:mm:ss.fff"));
            Status.Add("Fail");
            link = sclink;
            lineNo = lineno;
            Message = Msg;
            EXType = type;

        }
        public void IterationEndFunction()
        {
            for (int p = 0; p <= Status.Count - 1; p++)
            {
                if (Status[p].Equals("Pass"))
                {
                    passed++;
                }
                else if (Status[p].Equals("Fail"))
                {
                    failed++;
                }
            }
            starttimecount = StartTime.Count;

            passedsteps = passed;
            failedsteps = failed;


        }
        public void OverAllEndTime()
        {
            OverallEndTime = System.DateTime.Now.ToString("hh:mm:ss.fff");
        }

        public void htmlwriter()
        {
            Iteration = new StringBuilder();

            for (int m = 1; m <= RowCount; m++)
            {
                string path = string.Format("\"{0}\"", folderpath + "\\Iteration_" + m + ".html");
                var tab = "<tr><td><a href = " + path + ">Iteration " + m + "</a></td></tr>";
                Iterationtable.Add(tab);
            }
            foreach (string id in Iterationtable)
            {
                Iteration.Append(id + " ");
            }

            Rows = new StringBuilder();

            IterationStartTime = OverallStartTime;
            IterationEndTime = OverallEndTime;
            Iterationdue = Convert.ToDateTime(IterationEndTime) - Convert.ToDateTime(IterationStartTime);
            Iterationduration = Iterationdue.ToString().Substring(0, 12);

            for (int j = 0; j <= StartTime.Count - 1; j++)
            {
                var mn = MethodName[j];
                var st = StartTime[j];
                var et = EndTime[j]; // System.DateTime.Now.ToString("HH:MM:SS.FFF");                
                duration = Convert.ToDateTime(et) - Convert.ToDateTime(st);
                string duration1 = duration.ToString();
                string duration2 = duration1.Substring(0, 12);

                string c = string.Format("\"{0}\"", "td1");
                var table1 = "<tr><td id=" + c + "> " + MethodName[j] + "</td><td>" + StartTime[j] + "</td><td>" + EndTime[j] + "</td><td>" + duration2 + "</td><td>" + Status[j] + "</td></tr>";
                Table.Add(table1);
            }
            foreach (string id in Table)
            {
                Rows.Append(id + " ");
            }

            if (g1 == RowCount)
            {
                ExecutionEndTime = System.DateTime.Now.ToString("hh:mm:ss.fff");
                ExecutionEndTime1 = System.DateTime.Now.ToString("HH:mm:ss.fff");
            }
        }

        public string IterationResultWriter()
        {
            //ReportLocation = Reportlocation;
            //Iteration Result Writer
            mplocation = folderpath + "\\MainPage.html";
            string body1 = string.Empty;
            StreamReader reader1 = new StreamReader("\\\\192.169.1.55\\Automation reports\\SCUF\\Report Repository\\Templates\\HTML Template\\IterationPage_Template.html");
            body1 = reader1.ReadToEnd();
            body1 = body1.Replace("{MainPageLocation}", mplocation);
            body1 = body1.Replace("{username}", UserName);
            body1 = body1.Replace("{platform}", OSName);
            body1 = body1.Replace("{browsername}", BrowserName);
            body1 = body1.Replace("{Scenario}", Scenario);
            body1 = body1.Replace("{Iteration Start Time}", IterationStartTime);
            body1 = body1.Replace("{Iteration End Time}", IterationEndTime);
            body1 = body1.Replace("{Duration}", Iterationduration);
            body1 = body1.Replace("{No}", g1.ToString());
            body1 = body1.Replace("{Iteration}", Iteration.ToString());
            body1 = body1.Replace("{rows}", Rows.ToString());
            body1 = body1.Replace("{Screenshot}", link);
            double PassedStep = (double)passedsteps / totalsteps;
            double PassedStep1 = PassedStep * 180;
            body1 = body1.Replace("{PassedSteps}", passedsteps.ToString());
            body1 = body1.Replace("{Passedsteps}", PassedStep1.ToString());
            body1 = body1.Replace("{Totalsteps}", totalsteps.ToString());
            int stepsNE = totalsteps - passedsteps;
            int stepsNE1 = totalsteps - passedsteps - failedsteps;
            double FailedStep = (double)stepsNE / totalsteps;
            double FailedStep1 = FailedStep * 180;
            body1 = body1.Replace("{StepsNotExecuted}", stepsNE.ToString());
            body1 = body1.Replace("{Failedsteps}", FailedStep1.ToString());
            double steppasspercentage = (double)passedsteps / totalsteps;
            double steppasspercentage1 = (double)steppasspercentage * 100;
            long sp1 = Convert.ToInt64(steppasspercentage1);
            body1 = body1.Replace("{StepWisePassPercentage}", sp1.ToString());

            if (failedsteps != 0)
            {
                Feildset = "<fieldset><legend>Log Info: </legend><br>Log Written Date:<br>" + System.DateTime.Now.ToString() + "<br><br>Error Line No:<br>" + lineNo + "<br><br>Exception Type:<br>" + EXType + "<br><br>Error Msg:<br>" + Message + "</fieldset>";
                body1 = body1.Replace("{feildset}", Feildset);
            }
            else
            {
                body1 = body1.Replace("{feildset}", "");
            }

            reader1.Close();
            filelocation3 = folderpath + "\\Iteration_" + g1 + ".html";
            StreamWriter sw13 = new StreamWriter(filelocation3);
            sw13.WriteLine(body1);
            sw13.Close();

            StartTime.Clear();
            EndTime.Clear();
            Status.Clear();
            MethodName.Clear();
            Rows.Clear();
            Table.Clear();
            passed = 0;
            failed = 0;

            return body1;
        }
    }
    public partial class Result : Report
    {
        private static ArrayList IterationList = new ArrayList();
        private static ArrayList ScenarioList = new ArrayList();
        private static ArrayList TestCaseIDList = new ArrayList();
        private static ArrayList ExpectedStatusList = new ArrayList();
        private static ArrayList IterationStartTimeList = new ArrayList();
        private static ArrayList IterationEndTimeList = new ArrayList();
        private static ArrayList IterationDurationList = new ArrayList();
        private static ArrayList TotalPassedStepsList = new ArrayList();
        private static ArrayList TotalFailedStepsList = new ArrayList();
        private static ArrayList StepsNotExecuted = new ArrayList();
        private static ArrayList IterationStatus = new ArrayList();
        private static ArrayList FinalResultList = new ArrayList();
        //private static ArrayList PassedStepsCount = new ArrayList();
        //private static ArrayList FailedStepsCount = new ArrayList();
        private static ArrayList TestCaseDescriptionList = new ArrayList();
        private static ArrayList ExpectedResultedList = new ArrayList();
        private static ArrayList filelocation3List = new ArrayList();
        public List<string> ResultTable = new List<string>();
        public List<string> Barchart = new List<string>();
        public List<string> MainPageIterationtable = new List<string>();
        public StringBuilder Maintable, MainPageIteration, ChartTable;
        public TimeSpan IterartionExecutionDuration;
        public string FinalResult, tab;
        private static int PassedIterartion = 0, FailedIteration = 0;
        private static ArrayList arr1 = new ArrayList();
        public StringBuilder Table1;
        public string Reportlocation, loglocation;
        private static ArrayList arr = new ArrayList();
        public StringBuilder Table;
        //private string arcolour, ecolour, acolour;
        public static int a2, sab;
        private static ArrayList arr2 = new ArrayList();
        public StringBuilder Table2;
        public static string re, re1, re2, re3, re4, re5, b5, b4, b6, b7;

        public void Final(int rowcount, string ReportLocation)
        {

            Reportlocation = ReportLocation;
            loglocation = Reportlocation + "\\logfile.txt";
            IterationList.Add(Report.g1);
            ScenarioList.Add(Report.Scenario);
            TestCaseIDList.Add(Report.TestCaseID);
            ExpectedStatusList.Add(Report.ExpectedStatus);
            IterationStartTimeList.Add(Report.OverallStartTime);
            IterationEndTimeList.Add(Report.OverallEndTime);
            IterationDurationList.Add(Report.Iterationduration);
            TotalPassedStepsList.Add(Report.passedsteps);
            TotalFailedStepsList.Add(Report.failedsteps);
            StepsNotExecuted.Add(Report.totalsteps - Report.passedsteps - Report.failedsteps);
            TestCaseDescriptionList.Add(Report.TestCaseDescription);
            ExpectedResultedList.Add(Report.ExpectedResulted);
            filelocation3List.Add(Report.filelocation3);
            //PassedStepsCount.Add()

            if (Report.failedsteps == 0)
            {
                IterationStatus.Add("Pass");
            }
            else
            {
                IterationStatus.Add("Fail");
            }

            Maintable = new StringBuilder();
            ChartTable = new StringBuilder();

            for (int j = 0; j <= Report.g1 - 1; j++)
            {
                var Iteration = IterationList[j];
                var Scenario = ScenarioList[j];
                var Testcaseid = TestCaseIDList[j];
                var expectedstatus = ExpectedStatusList[j];
                var IterationStartTime = IterationStartTimeList[j];
                var IterationEndTime = IterationEndTimeList[j];
                var IterationDuration = IterationDurationList[j];
                var TotalPassedSteps = TotalPassedStepsList[j];
                var TotalFailedSteps = TotalFailedStepsList[j];
                var StepsNotexecuted = StepsNotExecuted[j];
                var Filelocation = filelocation3List[j];

                ActualResult = IterationStatus[j].ToString();

                if (ActualResult == expectedstatus.ToString())
                {
                    FinalResult = "Pass";
                }
                else
                {
                    FinalResult = "Fail";
                }
                string loc = string.Format("\"{0}\"", Filelocation);

                var table1 = "<tr href=" + loc + "><td>" + Iteration + "</td><td>Iteration" + Iteration + "</td><td>"+Testcaseid+"</td><td> " + Scenario + "</td><td>" + IterationStartTime + "</td><td>" + IterationEndTime + "</td><td>" + IterationDuration + "</td><td>" + ActualResult.ToString() + "</td><td>" + expectedstatus.ToString() + "</td><td>" + FinalResult + "</td></tr>";
                ResultTable.Add(table1);


                //-----------------------------------------------------------------------------

                if (Report.RowCount == Report.g1)
                {
                    FinalResultList.Add(FinalResult);

                }

            }


            foreach (string id in ResultTable)
            {
                Maintable.Append(id + " ");
            }
            
        }

        public string MainPageWriter()
        {
            //Main Page Writer
            string body = string.Empty;
            StreamReader reader = new StreamReader("\\\\192.169.1.55\\Automation reports\\SCUF\\Report Repository\\Templates\\HTML Template\\OverallResullt_Template.html");
            body = reader.ReadToEnd();
            body = body.Replace("{username}", Report.UserName);
            body = body.Replace("{platform}", Report.OSName);
            body = body.Replace("{URL}", Report.LURL);
            body = body.Replace("{Application}", Report.ApplicationName);
            body = body.Replace("{browsername}", Report.BrowserName.ToUpper());
            body = body.Replace("{ProjectName}", Report.ProjName);
            //body = body.Replace("{ExecutionStartTime}", Class1.ExecutionStartTime);
            //body = body.Replace("{ExecutionEndTime}", Class1.ExecutionEndTime);
            IterartionExecutionDuration = Convert.ToDateTime(Report.ExecutionEndTime1) - Convert.ToDateTime(Report.ExecutionStartTime1);
            string duration = IterartionExecutionDuration.ToString().Substring(0, 12);
            body = body.Replace("{Duration}", duration);

            for (int r = 0; r <= Report.g1 - 1; r++)
            {
                if (FinalResultList[r].ToString() == "Pass") { PassedIterartion++; }
                if (FinalResultList[r].ToString() == "Fail") { FailedIteration++; }
            }

            body = body.Replace("{Iterations}", Report.g1.ToString());
            body = body.Replace("{Passed}", PassedIterartion.ToString());
            body = body.Replace("{Failed}", FailedIteration.ToString());
            double PassedPercent = (double)PassedIterartion / Report.g1;
            double PassedPercent1 = PassedPercent * 180;
            double FailedPercent = (double)FailedIteration / Report.g1;
            double FailedPercent1 = FailedPercent * 180;
            body = body.Replace("{PassedPercent}", PassedPercent1.ToString());
            body = body.Replace("{FailedPercent}", FailedPercent1.ToString());
            decimal passed = System.Convert.ToDecimal(PassedIterartion);
            decimal total = System.Convert.ToDecimal(Report.RowCount);
            decimal percentage = (passed / total) * 100;
            Percentage = System.Convert.ToInt16(percentage);
            body = body.Replace("{Percentage}", Percentage.ToString());
            body = body.Replace("{Maintable}", Maintable.ToString());


            reader.Close();
            string filelocation2 = Report.mplocation;
            StreamWriter sw1 = new StreamWriter(filelocation2);
            sw1.WriteLine(body);
            sw1.Close();
            return body;
        }

        public string MainPageWriter1()
        {
            //Main Page Writer
            string body1 = string.Empty;
            StreamReader reader = new StreamReader("\\\\192.169.1.55\\Automation reports\\SCUF\\Report Repository\\Templates\\HTML Template\\MainPage_Template.html");
            body1 = reader.ReadToEnd();
            body1 = body1.Replace("{username}", Report.UserName);
            body1 = body1.Replace("{platform}", Report.OSName);
            body1 = body1.Replace("{URL}", Report.LURL);
            body1 = body1.Replace("{Application}", Report.ApplicationName);
            body1 = body1.Replace("{browsername}", Report.BrowserName.ToUpper());
            body1 = body1.Replace("{ProjectName}", Report.ProjName);

            for (int r = 0; r <= Report.g1 - 1; r++)
            {
                if (FinalResultList[r].ToString() == "Passed") { PassedIterartion++; }
                if (FinalResultList[r].ToString() == "Failed") { FailedIteration++; }
            }

            body1 = body1.Replace("{Iterations}", Report.g1.ToString());
            body1 = body1.Replace("{Passed}", PassedIterartion.ToString());
            body1 = body1.Replace("{Failed}", FailedIteration.ToString());
            double PassedPercent = (double)PassedIterartion / Report.g1;
            double PassedPercent1 = PassedPercent * 180;
            double FailedPercent = (double)FailedIteration / Report.g1;
            double FailedPercent1 = FailedPercent * 180;
            body1 = body1.Replace("{PassedPercent}", PassedPercent1.ToString());
            body1 = body1.Replace("{FailedPercent}", FailedPercent1.ToString());
            decimal passed = System.Convert.ToDecimal(PassedIterartion);
            decimal total = System.Convert.ToDecimal(Report.RowCount);
            decimal percentage = (passed / total) * 100;
            Percentage = System.Convert.ToInt16(percentage);
            body1 = body1.Replace("{Percentage}", Percentage.ToString());
            //body = body.Replace("{Maintable}", Maintable.ToString());

            for (int j = 0; j <= Report.RowCount - 1; j++)
            {

                var TestCaseDescription = TestCaseDescriptionList[j].ToString();
                var ExpectedStatus = ExpectedStatusList[j].ToString();
                var ExpectedResulted = ExpectedResultedList[j].ToString();
                var FinalResult = FinalResultList[j].ToString();
                ActualResult = IterationStatus[j].ToString();
                var Filelocation = filelocation3List[j].ToString();

                body1 = body1.Replace("{TestDescription}", TestCaseDescription);
                body1 = body1.Replace("{ExpectedResult}", ExpectedResulted);
                body1 = body1.Replace("{SlNo}", (j + 1).ToString());

                string arcolour = "";
                string ecolour = "";
                string acolour = "";

                if (FinalResult == "Pass")
                {
                    acolour = "DARKGREEN";
                }
                else
                {
                    acolour = "MAROON";
                }

                if ((ExpectedStatus == "Pass"))
                {
                    ecolour = "DARKGREEN";
                }
                else
                {
                    ecolour = "MAROON";
                }

                if ((ActualResult.ToString() == "Pass"))
                {
                    arcolour = "DARKGREEN";
                }
                else
                {
                    arcolour = "MAROON";
                }
                string close = ">";


                string Filelocation1 = Regex.Replace(Filelocation, @"\\", @"\\"); ;
                string loc = string.Format("\"{0}\"", "location.href='" + Filelocation1 + "'");

                //-----------------------For div1--------------------//
                string st = "<div onclick=";
                string string1 = " style=";
                string string2 = string.Format("\"{0}\"", "height:140px; width:950px; margin-left: 15px; margin-top: 7px; border: 3px solid black; float : left; box-shadow: 5px 3px 3px 3px;");
                string s = st + loc + string1 + string2 + close;
                Console.WriteLine(s);

                //-----------------------For div2--------------------//
                string string3 = "<div style=";
                string string4 = string.Format("\"{0}\"", "height:135px; width:110px; border: 1.5px solid white; float : left;");
                //close
                string string5 = "<h1>" + (j + 1).ToString() + "<h1></div>";
                string s1 = string3 + string4 + close + string5;
                Console.WriteLine(s1);

                //-----------------------For div3--------------------//
                string string6 = "<div style=";
                string string7 = string.Format("\"{0}\"", "height:100px; width:833px; border: 1.5px solid white; float : left");
                //Close
                string s2 = string6 + string7 + close;
                Console.WriteLine(s2);

                //-----------------------For Table--------------------//
                string string8 = "<table id = ";
                string string9 = string.Format("\"{0}\"", "mytable");
                string string10 = "align=";
                string string11 = string.Format("\"{0}\"", "left");
                string string12 = "cellpadding=";
                string string13 = string.Format("\"{0}\"", "1");
                string string14 = "style =";
                string string15 = string.Format("\"{0}\"", "width:825; height:90; border collapse: 1px solid gray; margin-left: 3; text-align: center; color: white;");
                //close
                string s3 = string8 + string9 + string10 + string11 + string12 + string13 + string14 + string15 + close;
                Console.WriteLine(s3);

                //-----------------------For Col--------------------//
                string string16 = "<col width=";
                string string17 = string.Format("\"{0}\"", "169");
                //Close
                string string18 = "<col width=";
                string string19 = string.Format("\"{0}\"", "auto");
                //Close
                string s4 = string16 + string17 + close + string18 + string19 + close;
                Console.WriteLine(s4);

                //-----------------------For start tbody--------------------//
                string string20 = "<tbody>";
                string s5 = string20;
                Console.WriteLine(s5);

                //-----------------------For 1st tr/td--------------------//
                string string21 = "<tr><td><b>Test Case Description</b></td><td>:</td><td>" + TestCaseDescription + "</td></tr>";
                string s6 = string21;
                Console.WriteLine(s6);

                //-----------------------For 2nd tr/td--------------------//
                string string22 = "<tr><td><b>Expected/Resulted</b></td><td>:</td><td>" + ExpectedResulted + "</td></tr>";
                string s7 = string22;
                Console.WriteLine(s7);

                //-----------------------For end tbody--------------------//
                string string23 = "</tbody>";
                string s8 = string23;
                Console.WriteLine(s8);

                //-----------------------For end table--------------------//
                string string24 = "</table>";
                string s9 = string24;
                Console.WriteLine(s9);

                //-----------------------For end div--------------------//
                string string25 = "</div>";
                string s10 = string25;
                Console.WriteLine(s10);

                //-----------------------For div4--------------------//
                string string26 = "<div style=";
                string string27 = string.Format("\"{0}\"", "height:35px; width:833px; border: 0px solid white; float : left");
                //Close
                string s11 = string26 + string27 + close;
                Console.WriteLine(s11);

                //-----------------------For div5--------------------//
                string string28 = "<div style=";
                string string29 = string.Format("\"{0}\"", "height:35px; width:277.5px; border: 1.5px solid white; float : left; background:" + arcolour + ";");
                //Close
                string string30 = "<h2>Automation Result : " + ActualResult + "</h2></div>";
                string s12 = string28 + string29 + close + string30;
                Console.WriteLine(s12);

                //-----------------------For div6--------------------//
                string string31 = "<div style=";
                string string32 = string.Format("\"{0}\"", "height:35px; width:277.5px; border: 1.5px solid white; float : left; background:" + ecolour + ";");
                //Close
                string string33 = "<h2>Expected Result : " + ExpectedStatus + "</h2></div>";
                string s13 = string31 + string32 + close + string33;
                Console.WriteLine(s13);

                //-----------------------For div6--------------------//
                string string34 = "<div style=";
                string string35 = string.Format("\"{0}\"", "height:35px; width:277.5px; border: 1.5px solid white; float : left; background:" + acolour + ";");
                //Close
                string string36 = "<h2>Final Result : " + FinalResult + "</h2></div>";
                string s14 = string34 + string35 + close + string36;
                Console.WriteLine(s14);

                //-----------------------For end div--------------------//
                string string37 = "</div>";
                string s15 = string37;
                Console.WriteLine(s15);

                //-----------------------For end div--------------------//
                string string38 = "</div>";
                string s16 = string38;
                Console.WriteLine(s16);

                arr.Add(s);
                arr.Add(s1);
                arr.Add(s2);
                arr.Add(s3);
                arr.Add(s4);
                arr.Add(s5);
                arr.Add(s6);
                arr.Add(s7);
                arr.Add(s8);
                arr.Add(s9);
                arr.Add(s10);
                arr.Add(s11);
                arr.Add(s12);
                arr.Add(s13);
                arr.Add(s14);
                arr.Add(s15);
                arr.Add(s16);

                if (Report.RowCount - 1 == j)
                {
                    Table = new StringBuilder();
                    foreach (string id in arr)
                    {
                        Table.Append(id + "");
                    }

                    body1 = body1.Replace("{Table}", Table.ToString());

                    ArrayList Date = new ArrayList();
                    ArrayList Passed = new ArrayList();
                    ArrayList Failed = new ArrayList();
                    ArrayList PassPercent = new ArrayList();
                    ArrayList TotalIteration = new ArrayList();
                    ArrayList passed1 = new ArrayList();
                    ArrayList failed1 = new ArrayList();
                    ArrayList pp = new ArrayList();
                    ArrayList ti = new ArrayList();

                    a2 = 0;
                    
                    string[] tr = File.ReadAllLines(loglocation);
                    ArrayList a = new ArrayList(tr);
                    int a1 = a.Count;
                    Thread.Sleep(1000);
                    for (int i = a1; i <= a1; i--)
                    {
                        int i1 = i;
                        Date.Add(a[i1 - 5]);
                        PassPercent.Add(a[i1 - 4]);
                        TotalIteration.Add(a[i1 - 3]);
                        Passed.Add(a[i1 - 2]);
                        Failed.Add(a[i1 - 1]);

                        i = i - 4;
                        if (i == 1) { break; }
                    }

                    for (int sa = 0; sa <= Passed.Count - 1; sa++)
                    {
                        //-------------------------------------Pass Percent----------------------//

                        string[] res = Regex.Split(PassPercent[sa].ToString(), "-");
                        foreach (string res1 in res)
                        {
                            Console.WriteLine(res1);
                            pp.Add(res1);
                        }

                        //-------------------------------------Total Iteration----------------------//
                        string[] toti = Regex.Split(TotalIteration[sa].ToString(), "-");
                        foreach (string toti1 in toti)
                        {
                            Console.WriteLine(toti1);
                            ti.Add(toti1);
                        }

                        //-------------------------------------Passed----------------------//
                        string[] passs = Regex.Split(Passed[sa].ToString(), "-");
                        foreach (string passs1 in passs)
                        {
                            Console.WriteLine(passs1);
                            passed1.Add(passs1);
                        }

                        //-------------------------------------Failed----------------------//
                        string[] failll = Regex.Split(Failed[sa].ToString(), "-");
                        foreach (string failll1 in failll)
                        {
                            Console.WriteLine(failll1);
                            failed1.Add(failll1);

                        }

                        a2++;

                    }

                    string close1 = ">";
                    string cont = "<div class=";
                    string cont1 = string.Format("\"{0}\"", "container");
                    string c = cont + cont1 + close1;

                    sab = 1;
                    for (int sb = 0; sb <= PassPercent.Count - 1; sb++)
                    {
                        //-----------------------For h1--------------------//
                        string restring1 = string.Format("\"{0}\"", "Head2");
                        string restring2 = "<h2 class =" + restring1 + ">" + Date[sb] + "</h2>";
                        re = restring2;
                        Console.WriteLine(re);

                        //-----------------------For div2--------------------//
                        string restring3 = c + cont + string.Format("\"{0}\"", "skills");
                        // string string4 =  string3;
                        string restring5 = "style=";
                        string restring6 = string.Format("\"{0}\"", "width: " + pp[sb + sab].ToString() + "; background-color: #32CD32;");
                        //Close
                        string restring7 = PassPercent[sb] + "</div></div>";
                        re2 = restring3 + restring5 + restring6 + close1 + restring7;
                        Console.WriteLine(re2);

                        string restring61 = string.Format("\"{0}\"", "width:100%; background-color: #008080;");
                        string restring71 = TotalIteration[sb] + "</div></div>";
                        re3 = restring3 + restring5 + restring61 + close1 + restring71;
                        Console.WriteLine(re3);

                        int Pass_Iteration = Convert.ToInt32(passed1[sb + sab]);
                        int Total_Iteration = Convert.ToInt32(ti[sb + sab]);
                        int Fail_Iteration = Convert.ToInt32(failed1[sb + sab]);

                        double totalpass = (double)Pass_Iteration / Total_Iteration;
                        double tp = totalpass * 100;
                        double totalfail = (double)Fail_Iteration / Total_Iteration;
                        double tf = totalfail * 100;

                        string restring62 = string.Format("\"{0}\"", "width: " + tp.ToString() + "%; background-color: #2196F3;");
                        string restring72 = Passed[sb] + "</div></div>";
                        re4 = restring3 + restring5 + restring62 + close1 + restring72;
                        Console.WriteLine(re4);

                        string restring63 = string.Format("\"{0}\"", "width: " + tf.ToString() + "%; background-color: #f44336;");
                        string restring73 = Failed[sb] + "</div></div>";
                        re5 = restring3 + restring5 + restring63 + close1 + restring73;
                        Console.WriteLine(re5);

                        sab++;

                        arr2.Add(re);
                        arr2.Add(re2);
                        arr2.Add(re3);
                        arr2.Add(re4);
                        arr2.Add(re5);
                    }

                    Table2 = new StringBuilder();
                    foreach (string id in arr2)
                    {
                        Table2.Append(id + "");
                    }

                    body1 = body1.Replace("{Table2}", Table2.ToString());

                     
                    reader.Close();
                    StreamWriter sw13 = new StreamWriter(folderpath+ "\\OverallStatus.html");
                    sw13.WriteLine(body1);
                    sw13.Close();

                }
            }

            return body1;
        }

        public void ExecutionLogWriter()
        {
            string strLogText = Percentage.ToString();
            string strLogText1 = Report.g1.ToString();
            string strLogText2 = PassedIterartion.ToString();
            string strLogText3 = FailedIteration.ToString();

            // Create a writer and open the file:
            StreamWriter log;
            
            if (!File.Exists(loglocation))
            {
                log = new StreamWriter(loglocation);
            }
            else
            {
                log = File.AppendText(loglocation);
            }
            // Write to the file:
            log.WriteLine("Execution Log - " + System.DateTime.Now.ToString("dd-MM-yyyy"));
            log.WriteLine("Pass Percentage - " + strLogText + "%");
            log.WriteLine("Total Iteration - " + strLogText1);
            log.WriteLine("Passed Iteration - " + strLogText2);
            log.WriteLine("Failed Iteration - " + strLogText3);
            log.Close();
        }

        public string MailWritter(string MailTO, string MailCC, string MailSubject, string EScenario, string ReportLocation)
        {
            using (SmtpClient SmtpServer = new SmtpClient())
            {

                System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                SmtpServer.Credentials = new System.Net.NetworkCredential("ashokkumar.a@shriramvalue.in", "Welcome@123");
                SmtpServer.Port = 587;
                SmtpServer.Host = "smtp.rediffmailpro.com";
                mail = new System.Net.Mail.MailMessage();
                mail.From = new MailAddress("automation@shriramvalue.in");

                foreach (var address in MailTO.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mail.To.Add(address);
                }
                foreach (var address1 in MailCC.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    mail.CC.Add(address1);
                }
                mail.Subject = MailSubject;


                StreamReader reader = new StreamReader("\\\\192.169.1.55\\Automation reports\\SCUF\\Report Repository\\Templates\\HTML Template\\EmailReportTemplate.html");
                string body1 = string.Empty;
                body1 = reader.ReadToEnd();
                body1 = body1.Replace("{Application}", Report.ProjName);
                body1 = body1.Replace("{Url}", Report.LURL);
                body1 = body1.Replace("{g}", Report.g1.ToString());
                body1 = body1.Replace("{Scenario}", EScenario);
                body1 = body1.Replace("{Table}", arr2[1].ToString());
                body1 = body1.Replace("{Table1}", arr2[2].ToString());
                body1 = body1.Replace("{Table2}", arr2[3].ToString());
                body1 = body1.Replace("{Table3}", arr2[4].ToString());

                string mailpage = folderpath + "\\Mail.html";
                StreamWriter sw1 = new StreamWriter(mailpage);
                sw1.WriteLine(body1);
                sw1.Close();

                string htmlfilepath = mailpage;
                var htmlToImageConv = new NReco.ImageGenerator.HtmlToImageConverter();
                string imagefilepath = folderpath + "\\EMailReport.jpeg";
                htmlToImageConv.GenerateImageFromFile(htmlfilepath, NReco.ImageGenerator.ImageFormat.Jpeg, imagefilepath);


                LinkedResource LinkedImage = new LinkedResource(imagefilepath);
                LinkedImage.ContentId = "MyPic";
                LinkedImage.ContentType = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Image.Jpeg);
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString("<img src=cid:MyPic>", null, "text/html");
                htmlView.LinkedResources.Add(LinkedImage);
                mail.AlternateViews.Add(htmlView);
                SmtpClient smtp = new SmtpClient("111.111.111.111", 25);
                System.Net.Mail.Attachment oAttch = new System.Net.Mail.Attachment(mplocation);
                System.Net.Mail.Attachment oAttch1 = new System.Net.Mail.Attachment(folderpath+ "\\OverallStatus.html");

               mail.Attachments.Add(oAttch);
                mail.Attachments.Add(oAttch1);
                SmtpServer.Send(mail);

                reader.Close();
                return body1;

            }
        }
    }

    public partial class AutomationPortal : Report
    {
        public static DataTable Table_Data;
        public static SqlCommand command;
        private static readonly string tablename;
        private static readonly int rowcount;
        private static readonly string Pk_Id;

        private static string temp;
        public static SqlConnection con;

        public static void PortalInitialization(string project, string Script, string sheetName, string tablename, string url, int r1)
        {

            tablename = project + "." + Script + "_" + sheetName + " as a";
            command = null;
            DataSet dtset = new DataSet();
            SqlDataAdapter sda = null;
            int Count = 0, Pk_Id = 0, iteration = 0; r1 = 0;
            Table_Data = new DataTable();
            con = new SqlConnection("Data Source=192.169.1.181\\sql2012;Initial Catalog=Automation;;User ID=sa;Password=welcome3#");
            try
            {
                command = new SqlCommand("select ROW_NUMBER() over(order by Primary_ID) as row,* from " + Script + "_" + sheetName, con);
                sda = new SqlDataAdapter(command);
                sda.Fill(dtset);
                command = new SqlCommand("select COUNT_BIG(*) from " + Script + "_" + sheetName, con);
                con.Open();
                count = Convert.ToInt32(command.ExecuteScalar());
                con.Close();
                command = new SqlCommand("select top(1) URL_Address from " + Script + "_" + sheetName, con);
                con.Open();
                url = command.ExecuteScalar().ToString();
                con.Close();

            }
            catch (Exception) { }

            if (dtset.Tables.Count > 0)
            {
                tablename = Script + "_" + sheetName;
                Table_Data = dtset.Tables[0];
                r1 = count;
                command = new SqlCommand("Drop table " + Script + "_" + sheetName, con);
                con.Open();
                command.ExecuteNonQuery();
                con.Close();
            }
            else
            {
                command = new SqlCommand("[dbo].[ProjectLoad]", con);
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@mode", "Get_TableRecords");
                command.Parameters.AddWithValue("@script", Script);
                command.Parameters.AddWithValue("@project", project);
                command.Parameters.AddWithValue("@tablename", tablename);
                sda = new SqlDataAdapter(command);
                sda.Fill(dtset);
                Count = Convert.ToInt16(dtset.Tables[0].Rows[0][0]);
                Pk_Id = Convert.ToInt16(dtset.Tables[1].Rows[0][0]);
                iteration = Convert.ToInt16(dtset.Tables[1].Rows[0][1]);
                url = dtset.Tables[1].Rows[0][2].ToString();        //newly added
                Table_Data = dtset.Tables[2];
                r1 = iteration != 0 ? Count < iteration ? Count : iteration : Count;
            }

           
        }
        public static void PortalSQLConnectionClosePass()
        {
            try
            {
                command = new SqlCommand("update " + tablename + " set status = 'Success' where g =" + temp, con);
                con.Open();
                command.ExecuteNonQuery();
                con.Close();

            }
            catch (Exception)
            {
            }
        }

        public static void PortalSQLConnectionCloseEnd()
        {
            if (g1 == (rowcount))
            {
                //added for portal:
                ////////////////////////////
                command = new SqlCommand("update Execution_Details set status = 'Success' where Pk_Id =" + Pk_Id, con);
                if (con.State == ConnectionState.Closed)
                    con.Open();
                command.ExecuteNonQuery();
                con.Close();


            }
        }
    }
}
