using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.IO;

namespace CDRcompare
{
    public partial class FrmFormattedCompare : Form
    {
        public FrmFormattedCompare()
        {
            InitializeComponent();
            textBox12.Text = "0";
            textBox13.Text = "0";
        }

        private void btnCompare_Click(object sender, EventArgs e)
        {
            bool ifOnlyCompareDistNum = false;

            if (checkBox1.Checked == true)
                ifOnlyCompareDistNum = true;



            if (string.IsNullOrEmpty(comboBox1.Text) || string.IsNullOrEmpty(comboBox2.Text))
            {
                MessageBox.Show("Pls select format", "CDRcompare");
                return;
            }

            int timeRange=0;

            if (CanCovert(textBox3.Text, typeof(System.Int32)))
                timeRange = Convert.ToInt32(textBox3.Text);
            else
            {
                MessageBox.Show("Pls input correct time range", "CDRcompare");
                return;
            }

            //存右边哈希值的Dictionary
            Dictionary<int, List<int>> dForSearchListRight = new Dictionary<int, List<int>>();

            List<CDRItem> listLeft = new List<CDRItem>();


            btnCompare.Enabled = false;

            label6.Text = "Loading...";
            ultraActivityIndicator1.Start(true);

            try
            {
                if (comboBox1.Text=="WHC format")
                    listLeft = FormatLoad.importExcelUsingNetvantageWHC(textBox1.Text, Convert.ToInt32(textBox4.Text), Convert.ToInt32(textBox5.Text), Convert.ToInt32(textBox6.Text), Convert.ToInt32(textBox7.Text));
                if (comboBox1.Text == "ashan format")
                    listLeft = FormatLoad.importExcelUsingNetvantageAshan(textBox1.Text, Convert.ToInt32(textBox4.Text), Convert.ToInt32(textBox5.Text), Convert.ToInt32(textBox6.Text), Convert.ToInt32(textBox7.Text));
                if (comboBox1.Text == "Umobile format")
                    listLeft = FormatLoad.importExcelUsingNetvantageUmobile(textBox1.Text, Convert.ToInt32(textBox4.Text), Convert.ToInt32(textBox5.Text), Convert.ToInt32(textBox6.Text), Convert.ToInt32(textBox7.Text));

            }
            catch (Exception loade)
            {
                MessageBox.Show(loade.Message);
            }

            if (listLeft==null)
            {
                MessageBox.Show("error format");
                btnCompare.Enabled = true;
                return;
            }


            List<CDRItem> listRight = new List<CDRItem>();

            try
            {
                if (comboBox2.Text == "WHC format")
                    listRight = FormatLoad.importExcelUsingNetvantageWHC(textBox2.Text, Convert.ToInt32(textBox8.Text), Convert.ToInt32(textBox9.Text), Convert.ToInt32(textBox10.Text), Convert.ToInt32(textBox11.Text));
                if (comboBox2.Text == "ashan format")
                    listRight = FormatLoad.importExcelUsingNetvantageAshan(textBox2.Text, Convert.ToInt32(textBox8.Text), Convert.ToInt32(textBox9.Text), Convert.ToInt32(textBox10.Text), Convert.ToInt32(textBox11.Text));
                if (comboBox2.Text == "Umobile format")
                    listRight = FormatLoad.importExcelUsingNetvantageUmobile(textBox2.Text, Convert.ToInt32(textBox8.Text), Convert.ToInt32(textBox9.Text), Convert.ToInt32(textBox10.Text), Convert.ToInt32(textBox11.Text));

            }
            catch (Exception loade)
            {
                MessageBox.Show(loade.Message);
            }

            if (listRight == null)
            {
                btnCompare.Enabled = true;
                MessageBox.Show("error format");
                return;
            }

            //写 dForSearchListRight
            for (int ii = 0; ii < listRight.Count; ii++)
            {
                CDRItem ci = listRight[ii];

                string sRowText = "";
                int iHash = 0;


                if (ifOnlyCompareDistNum == false)
                {
                    sRowText = sRowText + ci.ani;
                    sRowText = sRowText + ci.dest;
                }
                else
                {
                    sRowText = ci.dest;

                }
                //MessageBox.Show(sRowText);
                iHash = sRowText.GetHashCode();

                if (dForSearchListRight.ContainsKey(iHash))
                {
                    dForSearchListRight[iHash].Add(ii);
                }
                else
                {
                    dForSearchListRight.Add(iHash, new List<int>());
                    dForSearchListRight[iHash].Add(ii);
                }
            }

            //--------------------------

            Application.DoEvents();
            label6.Text = "Comparing...";
            //ultraActivityIndicator1.Start(true);

            //hash Dictionary already ok ,now start search matching record,
            //first search record in left but not in right
            int iDeleteCount = 0;
            List<int> listToBeDeleteLeft = new List<int>();
            string filePath = "compare_result.csv";

            for (int ii = 0; ii < listLeft.Count; ii++)
            {
                string strExportSame = "";
                

                string sRowText = "";
                int iHash = 0;

                CDRItem ci = listLeft[ii];
                strExportSame = ci.ani.ToString() + "," + ci.dest.ToString() + "," + ci.start.ToString() + "," + ci.duration.ToString() + ",";

                if (ifOnlyCompareDistNum == false)
                {
                    sRowText = sRowText + ci.ani;
                    sRowText = sRowText + ci.dest;
                }
                else
                {
                    sRowText =ci.dest;
                }

                //MessageBox.Show(sRowText);
                iHash = sRowText.GetHashCode();
                if (dForSearchListRight.ContainsKey(iHash))
                {
                    double minDurationDiff = 0.0;
                    double minDiffDuration = 0.0;

                    string minDiffAnum = null;
                    string minDiffBnum = null;
                    DateTime minDiffConnecttime=DateTime.MinValue;

                    bool exactMatch = false;
                    bool mohuMatch = false;
                    int recordElement = 0;
                    

                    foreach (int element in dForSearchListRight[iHash])     //loop A
                    {
                        //是否通过时间和duration检查
                        //if ((compareDateTime(ci.start, listRight[element].start, timeRange) == true))
                        if ((compareDateTime(ci.start.AddHours(Convert.ToDouble(textBox12.Text)), listRight[element].start.AddHours(Convert.ToDouble(textBox13.Text)), timeRange) == true))
                        {
                            if (System.Math.Abs(ci.duration - listRight[element].duration) < 10.0)
                            {
                                strExportSame = strExportSame + listRight[element].ani.ToString() + "," + listRight[element].dest.ToString() + "," + listRight[element].start.ToString() + "," + listRight[element].duration.ToString();
                                //File.AppendAllText(filePath, strExportSame);
                                //File.AppendAllText(filePath, strExportSame);
                                listToBeDeleteLeft.Add(element);
                                //在dictionary的list删除此条找到的纪录
                                dForSearchListRight[iHash].Remove(element);

                                iDeleteCount++;

                                exactMatch = true;

                                break;     //break 是跳出loop A
                            }
                            //if ((System.Math.Abs(ci.duration - listRight[element].duration) > 10.0) && (System.Math.Abs(ci.duration - listRight[element].duration) < 20))      //时间差距大于10
                            //{
                            //    strExportSame = strExportSame + listRight[element].ani.ToString() + "," + listRight[element].dest.ToString() + "," + listRight[element].start.ToString() + "," + listRight[element].duration.ToString() + "," + "same a b num but duration diff GT 10";
                            //    //File.AppendAllText(filePath, strExportSame);
                            //    //File.AppendAllText(filePath, strExportSame);
                            //    listToBeDeleteLeft.Add(element);
                            //    //在dictionary的list删除此条找到的纪录
                            //    dForSearchListRight[iHash].Remove(element);

                            //    //iDeleteCount++;      这里不能加 ，因为不是完全相等的记录

                            //    break;
                            //}

                            if ((System.Math.Abs(ci.duration - listRight[element].duration) > 10.0)&&  (System.Math.Abs(ci.duration - listRight[element].duration) < 30.0))
                            {
                                if (minDurationDiff == 0.0)
                                {
                                    minDurationDiff = System.Math.Abs(ci.duration - listRight[element].duration);
                                    minDiffAnum = listRight[element].ani;
                                    minDiffBnum = listRight[element].dest;
                                    minDiffConnecttime = listRight[element].start;
                                    minDiffDuration = listRight[element].duration;

                                    recordElement = element;
                                    mohuMatch = true;
                                    continue;
                                }
   
                                if (System.Math.Abs(ci.duration - listRight[element].duration) < minDurationDiff)
                                {
                                    minDurationDiff = System.Math.Abs(ci.duration - listRight[element].duration);
                                    minDiffAnum = listRight[element].ani;
                                    minDiffBnum = listRight[element].dest;
                                    minDiffConnecttime = listRight[element].start;
                                    minDiffDuration = listRight[element].duration;


                                    recordElement = element;
                                    mohuMatch = true;
                                    continue;
                                }
                            }


                        }

                    }     //loop A
                    if ((!exactMatch) && (mohuMatch))
                    {
                        strExportSame = strExportSame + minDiffAnum + "," + minDiffBnum + "," + minDiffConnecttime.ToString() + "," + minDiffDuration.ToString() + "," + "same a b num but duration diff GT 10";
                        listToBeDeleteLeft.Add(recordElement);
                        dForSearchListRight[iHash].Remove(recordElement);
                    }

                    
                    strExportSame = strExportSame + "\r\n";
                    File.AppendAllText(filePath, strExportSame);

                }
                else
                {
                    strExportSame = strExportSame + "\r\n";
                    File.AppendAllText(filePath, strExportSame);
                }

            }          //left--->right search ok

            if (MessageBox.Show("Search finished \r\nfrom left to right find " + iDeleteCount.ToString() + "  same records\r\nDo you want to export right_only record? ", "CDRcompare", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                //delete right record
                listToBeDeleteLeft.Sort();
                listToBeDeleteLeft.Reverse();
                for (int ii=0;ii<listToBeDeleteLeft.Count;ii++)
                    listRight.RemoveAt(listToBeDeleteLeft[ii]);

                for (int ii = 0; ii < listRight.Count; ii++)
                {
                    string strExportRemain = "";
                    strExportRemain=",,,,"+listRight[ii].ani.ToString()+","+listRight[ii].dest.ToString()+","+listRight[ii].start.ToString()+","+listRight[ii].duration.ToString()+"\r\n";
                    File.AppendAllText(filePath, strExportRemain);
                }

                MessageBox.Show("Export finished");
                Application.DoEvents();
                btnCompare.Enabled = true;
                label6.Text = "Finished";
                ultraActivityIndicator1.Stop(true);


            }
            btnCompare.Enabled = true;
            label6.Text = "Finished";
            ultraActivityIndicator1.Stop(true);

        }
        //
        //
        private bool compareDateTime(object o1, object o2, int DateDiffInMin)
        {
            if ((o1 == null) || (o2 == null))
                return false;

            if (!(o1.GetType() == typeof(System.DateTime)) || !(o2.GetType() == typeof(System.DateTime)))
                return false;

            DateTime dtO2 = Convert.ToDateTime(o2);
            //对时间的判断 允许有误差 

            double timeDiff = dtO2.Subtract((DateTime)o1).TotalMinutes;

            if (timeDiff < 0)
                timeDiff = -timeDiff;

            if (timeDiff <= DateDiffInMin)
                return true;
            else
                return false;
        }
        //
        private void button1_Click(object sender, EventArgs e)
        {
            //get path
            OpenFileDialog op = new OpenFileDialog();
            op.RestoreDirectory = true;
            op.Filter = "All|*.*|xlsx|*.xlsx";
            if (op.ShowDialog() != DialogResult.OK)
                return;

            textBox1.Text = op.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //get path
            OpenFileDialog op = new OpenFileDialog();
            op.RestoreDirectory = true;
            op.Filter = "All|*.*|xlsx|*.xlsx";
            if (op.ShowDialog() != DialogResult.OK)
                return;

            textBox2.Text = op.FileName;
        }
        //
        private Boolean CanCovert(String value, Type type)
        {
            TypeConverter converter = TypeDescriptor.GetConverter(type);
            return converter.IsValid(value);
        }
        //
        private void FrmFormattedCompare_Load(object sender, EventArgs e)
        {
            XElement xmldoc = XElement.Load("formatnew.xml");

            var result = xmldoc.Descendants("formatName")
                              .ToList();
            foreach (string strl in result)
            {
                comboBox1.Items.Add(strl);
                comboBox2.Items.Add(strl);
            }

            textBox3.Text = Properties.Settings.Default.defaultrange.ToString();

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(comboBox1.SelectedItem.ToString());

            //从formatnew.xml查询 SrcNum,DistNum,ConnectTime,Duration的列索引 

            XElement xmldoc = XElement.Load("formatnew.xml");

            var ColCount = xmldoc.Descendants("format")
              .Where(p => (string)p.Element("formatName") == comboBox1.SelectedItem.ToString())
              .Descendants("ImportantColumn")
              .Select(p => p.Value)
              .ToList();

            List<int> liColumnNum = new List<int>();

            foreach (string strt in ColCount)
            {
                liColumnNum.Add(Convert.ToInt32(strt));

            }
            textBox4.Text = liColumnNum[0].ToString();
            textBox5.Text = liColumnNum[1].ToString();
            textBox6.Text = liColumnNum[2].ToString();
            textBox7.Text = liColumnNum[3].ToString();
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
           
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            XElement xmldoc = XElement.Load("formatnew.xml");

            var ColCount = xmldoc.Descendants("format")
              .Where(p => (string)p.Element("formatName") == comboBox2.SelectedItem.ToString())
              .Descendants("ImportantColumn")
              .Select(p => p.Value)
              .ToList();

            List<int> liColumnNum = new List<int>();

            foreach (string strt in ColCount)
            {
                liColumnNum.Add(Convert.ToInt32(strt));

            }
            textBox8.Text = liColumnNum[0].ToString();
            textBox9.Text = liColumnNum[1].ToString();
            textBox10.Text = liColumnNum[2].ToString();
            textBox11.Text = liColumnNum[3].ToString();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
