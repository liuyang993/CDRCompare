using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Infragistics.Documents.Excel;
using System.Windows.Forms;
using System.ComponentModel;
using System.Globalization;

namespace CDRcompare
{
    static class FormatLoad
    {
        public static List<CDRItem> importExcelUsingNetvantageWHC(string strFilePath,int iSrc,int iDist,int iConTime,int iDuration)
        {
            List<CDRItem> list = new List<CDRItem>();
            bool bOADateFormat = false;
            bool bDatetimeFormat = false;
            bool bXueliangFormat = false;

            //----------to use this function ,remember to add reference :using Infragistics.Documents.Excel;
            if (File.Exists(strFilePath))
            {
                //Load the Excel File into the Workbook Object
                //Application.DoEvents();
                Workbook theWorkbook = Workbook.Load(strFilePath, false);


                //We will only work with the first Worksheet in the Workbook      --   只支持第一个WorkSheet
                Worksheet theWorksheet = theWorkbook.Worksheets[0];

                int theRowCounter = 0;

                //Iterate through all Worksheet rows   Iterate through
                foreach (WorksheetRow theWorksheetRow in theWorksheet.Rows)
                {
                    if (theRowCounter == 0)
                    {
                        //check file format
                        /*
                        string strHeadArray="";
                        foreach (WorksheetCell theWorksheetCell in theWorksheetRow.Cells)
                        {
                            strHeadArray = strHeadArray + theWorksheetCell.Value.ToString().Trim() + ",";
                        }
                        strHeadArray = strHeadArray.Substring(0, strHeadArray.Length - 1);
                        //ID	CDRID	SRCNum	DSTNum	TGIN	IPIn	TGOUT	IPOut	CDRDate	SetupTime	ConnectTime	DisconnectTime	Duration	CauseValue	Country	City	Remarks
                        if (strHeadArray != "ID,CDRID,SRCNum,DSTNum,TGIN,IPIn,TGOUT,IPOut,CDRDate,SetupTime,ConnectTime,DisconnectTime,Duration,CauseValue,Country,City,Remarks")
                        {
                            return null;
                        }
                         * */

                        theRowCounter++;
                        continue;
                    }
                    else
                    //This is the actual data that will populate the data model
                    {
                        CDRItem ci = new CDRItem();
                        try
                        {
                            ci.ani = (theWorksheetRow.Cells[iSrc].Value != null) ? theWorksheetRow.Cells[iSrc].Value.ToString() : "";

                            ci.dest = (theWorksheetRow.Cells[iDist].Value != null) ? theWorksheetRow.Cells[iDist].Value.ToString() : "";


                            if ((bOADateFormat == false) && (bDatetimeFormat == false) && (bXueliangFormat == false))
                            {
                                try
                                {
                                    ci.start = DateTime.FromOADate(Convert.ToDouble(theWorksheetRow.Cells[iConTime].Value));
                                    //ci.start = DateTime.Parse(theWorksheetRow.Cells[iConTime].Value.ToString());
                                    bOADateFormat = true;


                                }
                                catch (Exception dep)
                                {
                                    ci.start = DateTime.Parse(theWorksheetRow.Cells[iConTime].Value.ToString());
                                    //ci.start = DateTime.FromOADate(Convert.ToDouble(theWorksheetRow.Cells[iConTime].Value));
                                    bDatetimeFormat = true;
                                }
                            }
                            else
                            {
                                if ((bOADateFormat == true) && (theWorksheetRow.Cells[iConTime].Value != null))
                                {
                                    ci.start = DateTime.FromOADate(Convert.ToDouble(theWorksheetRow.Cells[iConTime].Value));
                                }
                                if ((bDatetimeFormat == true) && (theWorksheetRow.Cells[iConTime].Value != null))
                                {
                                    ci.start = DateTime.Parse(theWorksheetRow.Cells[iConTime].Value.ToString());
                                }

                                if ((bXueliangFormat == true) && (theWorksheetRow.Cells[iConTime].Value != null))
                                {
                                    ci.start = DateTime.ParseExact(theWorksheetRow.Cells[iConTime].Value.ToString(), "yyyyMMddHHmmss", null);
                                }

                            }




                            ci.duration = (theWorksheetRow.Cells[iDuration].Value != null) ? Convert.ToDouble(theWorksheetRow.Cells[iDuration].Value) : 0.0;

                                //ci.duration = Convert.ToDouble(theWorksheetRow.Cells[iDuration].Value);
                          
                     

                        }
                        catch (Exception tep)
                        {
                            string sXueliangDatetime = "";
                            sXueliangDatetime = theWorksheetRow.Cells[iConTime].Value.ToString();

                            string formatString = "yyyyMMddHHmmss";

                            try
                            {
                                DateTime dt = DateTime.ParseExact(sXueliangDatetime, formatString, null);

                                ci.start = dt;

                                bXueliangFormat = true;
                            }
                            catch (Exception tmep)
                            { }


                            //MessageBox.Show(tep.Message + theRowCounter.ToString() + theWorksheetRow.Cells[0].Value.ToString() + theWorksheetRow.Cells[1].Value.ToString() + theWorksheetRow.Cells[2].Value.ToString() + theWorksheetRow.Cells[3].Value.ToString());
                        }

                        if (ci.dest != "")
                            list.Add(ci);
                     }

                    theRowCounter++;
                }

            }
            else
            {
                MessageBox.Show("No such file was found!");
            }         
            return list;
        }
        //
        public static List<CDRItem> importExcelUsingNetvantageAshan(string strFilePath, int iSrc, int iDist, int iConTime, int iDuration)
        {
            //要进行的处理 
            //1：去掉distNum前缀的70743     ok
            //2：srcNum如果等于asterisk，改为空   ok
            //3：srcNum如果等于数字，去掉前面多余的0，例如0080 改为80  ok
            //4：支持多个worksheet    ok

            List<CDRItem> list = new List<CDRItem>();
            //----------to use this function ,remember to add reference :using Infragistics.Documents.Excel;
            if (File.Exists(strFilePath))
            {
                //Load the Excel File into the Workbook Object
                //Application.DoEvents();
                Workbook theWorkbook = Workbook.Load(strFilePath, false);
                for (int jj = 0; jj < theWorkbook.Worksheets.Count; jj++)
                {
                    Worksheet theWorksheet = theWorkbook.Worksheets[jj];

                    int theRowCounter = 0;

                    //Iterate through all Worksheet rows   Iterate through
                    foreach (WorksheetRow theWorksheetRow in theWorksheet.Rows)
                    {
                        if (theRowCounter == 0)
                        {
                            //check file format
                            //string strHeadArray = "";
                            //foreach (WorksheetCell theWorksheetCell in theWorksheetRow.Cells)
                            //{
                            //    strHeadArray = strHeadArray + theWorksheetCell.Value.ToString().Trim() + ",";
                            //}
                            //strHeadArray = strHeadArray.Substring(0, strHeadArray.Length - 1);
                            

                            ////if (strHeadArray != "start time,duration,CalledIP,Callout callerID,Callout calledNum")
                            //if (!strHeadArray.Equals("start time,duration,CalledIP,Callout callerID,Callout calledNum", StringComparison.OrdinalIgnoreCase))
                            //{

                            //    return null;
                            //}

                            theRowCounter++;
                            continue;
                        }
                        else
                        //This is the actual data that will populate the data model
                        {
                            CDRItem ci = new CDRItem();
                            try
                            {
                                ci.ani = (theWorksheetRow.Cells[iSrc].Value != null) ? theWorksheetRow.Cells[iSrc].Value.ToString() : "";
                                if (ci.ani == "asterisk")
                                    ci.ani = "";

                                Int64 l1;

                                Int64.TryParse(ci.ani, out l1);

                                if (l1 != 0)
                                    ci.ani = l1.ToString();

                                ci.dest = (theWorksheetRow.Cells[iDist].Value != null) ? theWorksheetRow.Cells[iDist].Value.ToString() : "";
                                if (ci.dest.Length > 5)
                                    ci.dest = ci.dest.Substring(5);

                                if (theWorksheetRow.Cells[iConTime].Value.GetType() == typeof(double))
                                    ci.start = DateTime.FromOADate((double)theWorksheetRow.Cells[iConTime].Value);
                                else
                                    ci.start = DateTime.Parse(theWorksheetRow.Cells[iConTime].Value.ToString());

                                ci.duration = Convert.ToDouble(theWorksheetRow.Cells[iDuration].Value);
                            }
                            catch (System.OverflowException tep)
                            {
                                MessageBox.Show("stack over flow" + tep.Message);
                            }
                            list.Add(ci);
                        }

                        theRowCounter++;
                    }  //loop worksheet's row
                }//loop worksheet
            }
            else
            {
                MessageBox.Show("No such file was found!");
            }            
            return list;

        }
        //
        public static List<CDRItem> importExcelUsingNetvantageUmobile(string strFilePath, int iSrc, int iDist, int iConTime, int iDuration)
        {
            List<CDRItem> list = new List<CDRItem>();
            //----------to use this function ,remember to add reference :using Infragistics.Documents.Excel;
            if (File.Exists(strFilePath))
            {
                //Load the Excel File into the Workbook Object
                //Application.DoEvents();
                Workbook theWorkbook = Workbook.Load(strFilePath, false);

                //We will only work with the first Worksheet in the Workbook      --   只支持第一个WorkSheet
                Worksheet theWorksheet = theWorkbook.Worksheets[0];

                int theRowCounter = 0;

                //Iterate through all Worksheet rows   Iterate through
                foreach (WorksheetRow theWorksheetRow in theWorksheet.Rows)
                {
                    if (theRowCounter == 0)
                    {
                        //check file format
                        //string strHeadArray = "";
                        //foreach (WorksheetCell theWorksheetCell in theWorksheetRow.Cells)
                        //{
                        //    strHeadArray = strHeadArray + theWorksheetCell.Value.ToString().Trim() + ",";
                        //}
                        //strHeadArray = strHeadArray.Substring(0, strHeadArray.Length - 1);
                        ////event_start_date	event_start_time	anum	bnum	event_duration	billing_operator
                        //if (strHeadArray != "event_start_date,event_start_time,anum,bnum,event_duration,billing_operator")
                        //{
                        //    return null;
                        //}

                        theRowCounter++;
                        continue;
                    }
                    else
                    //This is the actual data that will populate the data model
                    {
                        CDRItem ci = new CDRItem();
                        
                        try
                        {
                            ci.ani = (theWorksheetRow.Cells[iSrc].Value != null) ? theWorksheetRow.Cells[iSrc].Value.ToString() : "";

                            ci.dest = (theWorksheetRow.Cells[iDist].Value != null) ? theWorksheetRow.Cells[iDist].Value.ToString() : "";
                            //ci.dest = "60" + ci.dest;
                            if (ci.dest.Length > 3)
                                ci.dest = ci.dest.Substring(3);

                            //取日期
                            if (theWorksheetRow.Cells[iConTime].Value.GetType() == typeof(double))
                                ci.start = DateTime.FromOADate((double)theWorksheetRow.Cells[iConTime].Value);
                            else
                            {
                                //ci.start = DateTime.ParseExact(theWorksheetRow.Cells[iConTime].Value.ToString(), "dd/MM/yy", new CultureInfo("en-US"));
                                //ci.start = DateTime.ParseExact(theWorksheetRow.Cells[iConTime].Value.ToString(), "dd/MM/yy", new CultureInfo("en-US"));
                                ci.start = DateTime.Parse(theWorksheetRow.Cells[iConTime].Value.ToString());
                                //年和日互换

                                int iyear = ci.start.Year;
                                int imonth = ci.start.Month;
                                int iday = ci.start.Day;

                                ci.start = new DateTime(iday+2000, imonth, iyear-2000);

                            }
                                //ci.start = DateTime.Parse(theWorksheetRow.Cells[iConTime].Value.ToString());

                            //取时间
                            
                            DateTime d2 = Convert.ToDateTime(theWorksheetRow.Cells[iConTime+1].Value);
                            TimeSpan timeSpan = d2.TimeOfDay;
                            ci.start = ci.start.Date.Add(timeSpan);

                            //if (theWorksheetRow.Cells[iConTime].Value.GetType() == typeof(double))
                            //    ci.start = ci.start.Date.Add(DateTime.FromOADate((double)theWorksheetRow.Cells[iConTime].Value).TimeOfDay);
                            //else
                            //    ci.start = ci.start.Date.Add(DateTime.Parse(theWorksheetRow.Cells[iConTime].Value.ToString()).TimeOfDay);

                              

                            ci.duration = Convert.ToDouble(theWorksheetRow.Cells[iDuration].Value);
                        }
                        catch (Exception tep)
                        {
                            MessageBox.Show(tep.Message + theRowCounter.ToString() + theWorksheetRow.Cells[0].Value.ToString() + theWorksheetRow.Cells[1].Value.ToString() + theWorksheetRow.Cells[2].Value.ToString() + theWorksheetRow.Cells[3].Value.ToString());
                        }
                        list.Add(ci);
                    }

                    theRowCounter++;
                }

            }
            else
            {
                MessageBox.Show("No such file was found!");
            }
            return list;
        }
        //
        static Boolean CanCovert(String value, Type type)
        {
            TypeConverter converter = TypeDescriptor.GetConverter(type);
            return converter.IsValid(value);
        }
        //

    }
}
