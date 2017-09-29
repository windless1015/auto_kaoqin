using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

using Excel = Microsoft.Office.Interop.Excel; 


namespace auto_kaoqin
{

    struct QingjiaOneStatus
    {
        string name;//请假人
        string qingjia_type;//请假类型: 年假,全薪病假,病假,事假,调休,年假,年度服务假,其他,产检假
        DateTime start_time;//请假起始时间
        DateTime end_time;//请假结束时间
        public double qingjia_span;//请假的有效时间间隔
        string user_qingjia_span;//用户填写的时间间隔
        public void JudgeValidSpan()
        {
            if (qingjia_type == "事假")
            {
                qingjia_span = GetQingjiaValidSpan(start_time, end_time);
            }
            else
            {
                qingjia_span = GetOtherHolidaySpan(start_time, end_time);
            }
        }

        public void SetData(string _name, string _qingjia_type, string _start_time, string _end_time,
            string _user_qingjia_span)
        {
            name = _name;
            qingjia_type = _qingjia_type;
            start_time = DateTime.Parse(_start_time);
            end_time = DateTime.Parse(_end_time);
            user_qingjia_span = _user_qingjia_span;
        }

        //事假,开始和结束日期都在一天之内
        private double WithinOneday(DateTime start_today, DateTime end_today)
        {
            string start_str = start_today.ToString("HH:mm");
            string end_str = end_today.ToString("HH:mm");
            double ts = 0;
            double start_double = start_today.Hour + start_today.Minute / 60.0, end_double = end_today.Hour + end_today.Minute / 60.0;
            //如果开始时间在9点之前
            if (start_str.CompareTo("09:00") < 0)
                start_double = 9;
            if (start_str.CompareTo("12:00") > 0 && start_str.CompareTo("13:30") < 0)
                start_double = 13.5;
            if (end_str.CompareTo("12:00") > 0 && end_str.CompareTo("13:30") < 0)
                end_double = 12;
            if (end_str.CompareTo("18:00") > 0)
                end_double = 18;

            //如果开始时间在上午,结束时间在下午
            if (start_double <= 12 && end_double <= 18 && end_double > 13.5)
            {
                ts = 12 - start_double;
                ts += end_double - 13.5;
            }
            else
                ts = end_double - start_double;
            return ts;
        }

        private double GetQingjiaValidSpan(DateTime start, DateTime end)
        {
            int N_day = (end - start).Days;

            double ts_total = 0; //总假期小时数
            if (N_day > 0) //如果不是同一天的
            {
                //先让start增加N天,需要单独申请一个变量add_datetime
                DateTime add_datetime = start.AddDays(N_day);

                double tmp_tspan = 7.5 * N_day;

                ts_total += (tmp_tspan);
                //如果增加N天后,时间还是小于end,说明 是 4.12 11:00 -- 4.14 10:00这种情况
                if (add_datetime < end)
                {
                    tmp_tspan = WithinOneday(add_datetime, end);
                    ts_total += tmp_tspan;
                }
                else
                {
                    tmp_tspan = WithinOneday(end, add_datetime);
                    ts_total -= tmp_tspan;
                }
            }
            else if (start.Date.Day != end.Date.Day)
            {
                ts_total = 8;
            }
            else  //同一天
            {
                ts_total = WithinOneday(start, end);
            }
            return ts_total;
        }

        //其他假期的计算算法
        private int WithinOtherHoliday(DateTime start, DateTime end, string start_str, string end_str)
        {
            int ts = 0;
            double start_double = start.Hour + start.Minute / 60.0, end_double = end.Hour + end.Minute / 60.0;
            //如果开始时间在9点之前
            if (start_str.CompareTo("09:00") < 0)
                start_double = 9;
            if (start_str.CompareTo("12:00") > 0 && start_str.CompareTo("13:30") < 0)
                start_double = 13.5;
            if (end_str.CompareTo("12:00") > 0 && end_str.CompareTo("13:30") < 0)
                end_double = 12;
            if (end_str.CompareTo("18:00") > 0)
                end_double = 18;

            if (start_double < 12.0 && end_double >= 13.5 && end_double <= 18.0) // 一个在上午一个在下午
                ts = 8;
            else if (start_double <= 13.5 && end_double <= 13.5) //都在上午
                ts = 4;
            else if (start_double >= 12.0 && start_double <= 18.0 && //都在下午
                end_double >= 12.0 && end_double <= 18.0)
                ts = 4;
            return ts;
        }

        private int GetOtherHolidaySpan(DateTime start, DateTime end)
        {
            string start_str = start.ToString("HH:mm");
            string end_str = end.ToString("HH:mm");

            int N_day = (end - start).Days;
            int ts_total = 0; //总假期小时数
            if (N_day > 0) //如果不是同一天的
            {
                ts_total += 8 * N_day;
                if (end_str.CompareTo("12:00") <= 0)
                    ts_total += 4;
                if (end_str.CompareTo("12:00") > 0 && end_str.CompareTo("23:59") < 0)
                    ts_total += 8;

            }
            else
                ts_total = WithinOtherHoliday(start, end, start_str, end_str);
            return ts_total;
        }

    };
    /// <summary>
    /// //////////////////一个用户一天的状态
    /// </summary>
    public struct OneUserDayStatus 
    {
       
       public int day_of_month;//1,2,3,4....,29,30
       public string user_name;
       public string status;
       public string extra_data;


        public void SetValue(int _day_of_month,string _user_name, string _status, string _extra_data)
        {
            day_of_month = _day_of_month;
            user_name = _user_name;
            status = _status;
            extra_data = _extra_data;
        }
    };

    public partial class main_Form : Form
    {
        private static string kaoqin_excel_path = "";
        private static string shenpi_excel_path = "";
        private static Excel.Application need_write_app = null;
        private static Excel.Workbook need_write_book = null;
        private static Excel.Worksheet need_write_sheet = null;//需要写入的sheet
        private static List<OneUserDayStatus> oneday_list = new List<OneUserDayStatus>();
        private static List<QingjiaOneStatus> qingjia_list = new List<QingjiaOneStatus>();
        private static List<int> week_end_list = new List<int>();
        private static int month_idx_whole = 0;//本次是几月
        private static object objlock = new object();

        public main_Form()
        {
            InitializeComponent();
        }



        //加载考勤excel数据
        private void btn_add_kaoqin_Click(object sender, EventArgs e)
        {
            //先判断起始日期是否符合逻辑, 起始日期 <=  结束日期
            if (dateTimePicker_start.Value.Date > dateTimePicker_end.Value.Date)
            {
                MessageBox.Show("起始日期需要小于等于结束日期", "错误");
                dateTimePicker_start.Value = dateTimePicker_end.Value;
                return;
            }
            //打开文件
            OpenFileDialog openFileDialog_kaoqin = new OpenFileDialog();
            openFileDialog_kaoqin.InitialDirectory = @"c:\User";
            openFileDialog_kaoqin.RestoreDirectory = true; //下次打开对话框是否定位到上次打开的目录
            openFileDialog_kaoqin.Filter = "考勤excle文件(*.xls或*.xlsx)|*.xls;*.xlsx|所有文件 (*.*)|*.*";
            openFileDialog_kaoqin.Title = "选择考勤数据文件";

            if (openFileDialog_kaoqin.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //点击ok按钮之后,把文件的路径写入状态栏里面
                this.toolStripStatusLabel1.Text = openFileDialog_kaoqin.FileName;
                kaoqin_excel_path = openFileDialog_kaoqin.FileName;

                //启动一个子线程来处理考勤数据
                ThreadStart Ts = new ThreadStart(ProcessKaoqinData);
                Thread Thread_kaoqin = new Thread(Ts);
                Thread_kaoqin.Start();
                this.toolStripStatusLabel1.Text = "软件正在处理考勤数据,请稍等!";
                Thread_kaoqin.IsBackground = true;
                //Thread_kaoqin.Join();//直到这个线程运行完毕才会执行后面的代码
                //Thread_kaoqin.Abort();
                //this.toolStripStatusLabel1.Text = "考勤数据已经处理完毕!";
            }
            
        }

        //加载审批excel数据
        private void btn_add_shenpi_Click(object sender, EventArgs e)
        {
            if (dateTimePicker_start.Value.Date > dateTimePicker_end.Value.Date)
            {
                MessageBox.Show("起始日期需要小于等于结束日期", "错误");
                dateTimePicker_end.Value = dateTimePicker_start.Value;
                return;
            }
            OpenFileDialog openFileDialog_shenpi = new OpenFileDialog();
            openFileDialog_shenpi.InitialDirectory = @"c:\User";
            openFileDialog_shenpi.RestoreDirectory = true; //下次打开对话框是否定位到上次打开的目录
            openFileDialog_shenpi.Filter = "审批excle文件(*.xls或*.xlsx)|*.xls;*.xlsx|所有文件 (*.*)|*.*";
            openFileDialog_shenpi.Title = "选择审批数据文件";

            if (openFileDialog_shenpi.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.toolStripStatusLabel1.Text = openFileDialog_shenpi.FileName;
                shenpi_excel_path = openFileDialog_shenpi.FileName;
                this.toolStripStatusLabel1.Text = "软件正在处理审批数据,请稍等!";
                //启动一个子线程来处理考勤数据
                ThreadStart Ts = new ThreadStart(ProcessShenpiData);
                Thread Thread_shenpi = new Thread(Ts);
                Thread_shenpi.Start();
                Thread_shenpi.IsBackground = true;
            }
        }

        private void TestInvoker(string text)
        {
            if (this.statusStrip1.InvokeRequired)
                this.statusStrip1.Invoke(
                    new MethodInvoker(() => this.toolStripStatusLabel1.Text = text));
            else this.toolStripStatusLabel1.Text = text;
        }

        private void main_Form_Load(object sender, EventArgs e)
        {
            //得到当前的日期,使得日期控件现实昨天的日期
            DateTime yesterday_date = DateTime.Now.AddDays(-1);
            dateTimePicker_start.Value = yesterday_date;
            dateTimePicker_end.Value = yesterday_date;


            //加载需要填充的最终excel文件
            OpenFileDialog openFileDialog_final = new OpenFileDialog();
            openFileDialog_final.InitialDirectory = @"c:\User";
            openFileDialog_final.RestoreDirectory = true; //下次打开对话框是否定位到上次打开的目录
            openFileDialog_final.Filter = "excle文件(*.xls或*.xlsx)|*.xls;*.xlsx|所有文件 (*.*)|*.*";
            openFileDialog_final.Title = "选择最终数据文件";

            while (openFileDialog_final.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                this.toolStripStatusLabel1.Text = openFileDialog_final.FileName;
                string final_excel = openFileDialog_final.FileName;

                object misValue = System.Reflection.Missing.Value;
                need_write_app = new Excel.Application();


                need_write_book = need_write_app.Workbooks.Open(final_excel, 0, true,
                    5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                need_write_sheet = (Excel.Worksheet)need_write_book.Worksheets.get_Item(1);

                Range curentCell = (Range)need_write_sheet.Cells[1, 1]; //只举例第一个单元格被合并
                string text = curentCell.Text; //单元格文本
                int year_idx = text.LastIndexOf("年");
                int month_idx = text.LastIndexOf("月");
                month_idx_whole = int.Parse(text.Substring(year_idx + 1, month_idx - year_idx - 1));


                //获取need_write_sheet的第4行的背景颜色
                string aaa = need_write_sheet.Cells[4, 37].Value;
                for (int i = 1; i <= 40; i++)
                {     
                    if( int.Parse(need_write_sheet.Cells[4, i].Interior.Color.ToString()) < 16777215)
                        week_end_list.Add(i);
                }

                break;//跳出 while循环
            }
    
        }

        

        //处理start_date到end_date的考勤数据
        private void ProcessKaoqinData()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(kaoqin_excel_path, 0, true,
                5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //MessageBox.Show(xlWorkSheet.get_Range("D2", "D2").Value2.ToString());
            //得到excel表格使用的有效行列
            int iRowsCount = xlWorkSheet.UsedRange.Cells.Rows.Count;
            string start_date_excle = "";//excle表格中最早的日期 2017-03-01
            string end_date_excle = "";//excel表格中最晚的日期
            string user_name = "";//第一个用户姓名后面用得到
            int valid_index = 0; //有效数据的起始行
            //找到第一个考勤号码大于等于10000的那一行,然后获取D列的日期
            GetFirstValidUserName(ref xlWorkSheet, out user_name, out valid_index, iRowsCount);
            GetEarlistLatestTime(ref xlWorkSheet, out start_date_excle, out end_date_excle, valid_index, iRowsCount, user_name);
            DateTime start_date_excle_datetime = Convert.ToDateTime(start_date_excle);
            DateTime end_date_excle_datetime = Convert.ToDateTime(end_date_excle);
            if (this.dateTimePicker_start.Value.Date < start_date_excle_datetime.Date)
            {
                string title = "excel表格中的最早日期为:" + start_date_excle_datetime.ToShortDateString() + ",最晚日期为:" +
                    end_date_excle_datetime.ToShortDateString();
                MessageBox.Show(title,"错误");
                //这个地方的代码需要修改.因为会产生跨线程访问ui控件的问题///
                //具体实现参考msdn: https://msdn.microsoft.com/zh-cn/library/ms171728(v=vs.110).aspx
                UpdateDateTimePickerStart(start_date_excle_datetime);
                return;
            }

            if (this.dateTimePicker_end.Value.Date > end_date_excle_datetime.Date)
            {
                string title = "excel表格中的最晚日期为:" + end_date_excle_datetime.ToShortDateString();
                MessageBox.Show(title, "错误");
                UpdateDateTimePickerEnd(end_date_excle_datetime.Date);
                return;
            }

            string temp_datetime_str = "";
            DateTime excel_string_datetime;
            Queue<DateTime> queue_one_user_record = new Queue<DateTime>();//一个人所有的考勤记录
            //然后从第一个有效考勤记录的人进行数据计算和分析
            for (int j = valid_index; j <= iRowsCount; j++)  //从有效行数开始进行分析
            {
                if (user_name == GetCellContent(j, 2, ref xlWorkSheet))
                {
                    temp_datetime_str = GetCellContent(j, 4, ref xlWorkSheet);//从excel表格中获取的时间字符串
                    excel_string_datetime = Convert.ToDateTime(temp_datetime_str);// 根据时间字符串 生成的时间对象
                    if (excel_string_datetime.Date >= this.dateTimePicker_start.Value.Date &&
                        excel_string_datetime.Date <= this.dateTimePicker_end.Value.Date) 
                    {
                        //从时间控件start日期 到结束日期进行循环
                        queue_one_user_record.Enqueue(excel_string_datetime);     
                    }
                    
                }
                else 
                {
                    
                    GetOneDayStackData(ref queue_one_user_record, user_name);
                    //把数据写入,先上锁
                    lock (objlock)
                    {
                        WriteToFinalExcel(ref oneday_list, user_name);
                    }
                    oneday_list.Clear();//清空list中所有元素
                    user_name = GetCellContent(j, 2, ref xlWorkSheet);//用户姓名重新赋值
                    queue_one_user_record.Clear();
                    j--;
                }


                if ( j == iRowsCount && queue_one_user_record.Count != 0)
                {
                    GetOneDayStackData(ref queue_one_user_record, user_name);
                    //把数据写入,先上锁
                    lock (objlock)
                    {
                        WriteToFinalExcel(ref oneday_list, user_name);
                    }
                    oneday_list.Clear();//清空list中所有元素
                    queue_one_user_record.Clear();
                }
            }
           
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit(); //这一句很重要,否则excel对象不能从内存中退出
            //释放资源
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            TestInvoker("考勤数据已经加载完毕");//更新状态条
        }


        //审批数据
        private void ProcessShenpiData()
        {
            Excel.Application xlShenpiApp = null;
            Excel.Workbook xlShenpiWorkBook = null;
            Excel.Worksheet xlShenpiWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;

            xlShenpiApp = new Excel.Application();
            xlShenpiWorkBook = xlShenpiApp.Workbooks.Open(kaoqin_excel_path, 0, true,
                5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            int sheet_num = xlShenpiWorkBook.Worksheets.Count; //得到审批excel的sheet数目

            //循环读取不同的sheet 
            for (int i = 1; i <= sheet_num; i++)
            {
                //读取当前的sheet
                xlShenpiWorkSheet = (Excel.Worksheet)xlShenpiWorkBook.Worksheets.get_Item(i);
                //获取本sheet的有效行列
                int this_sheet_row = xlShenpiWorkSheet.UsedRange.Cells.Rows.Count;
                int this_sheet_colunm = xlShenpiWorkSheet.UsedRange.Cells.Columns.Count;
                //遍历本sheet的所有行
                for (int j = 2; j <= this_sheet_row; j++)
                {
                    //第N行是 航盛19F消费金融 或者 航盛19F财富管理事业部 或者 集团公线&职能 进行判断
                    string position = GetCellContent(j, 14, ref xlShenpiWorkSheet);
                    if (position == "航盛19F消费金融" || position == "航盛19F财富管理事业部" || position == "集团公线&职能")
                    {
                        string name = GetCellContent(j, 8, ref xlShenpiWorkSheet);
                        string type = GetCellContent(j, 15, ref xlShenpiWorkSheet);
                        string start = GetCellContent(j, 16, ref xlShenpiWorkSheet);
                        string end = GetCellContent(j, 17, ref xlShenpiWorkSheet);
                        string self_pronunce = GetCellContent(j, 18, ref xlShenpiWorkSheet);
                        QingjiaOneStatus qingjia_one = new QingjiaOneStatus();
                        qingjia_one.SetData(name, type, start, end, self_pronunce);
                        qingjia_one.JudgeValidSpan();
                        qingjia_list.Add(qingjia_one);
                        Console.WriteLine("{0}, {1},时长:{2}, 自述:{3}", name, type, qingjia_one.qingjia_span, self_pronunce);
                    }
                }

            }

            xlShenpiWorkBook.Close(true, misValue, misValue);
            xlShenpiApp.Quit(); //这一句很重要,否则excel对象不能从内存中退出
            //释放资源
            releaseObject(xlShenpiWorkSheet);
            releaseObject(xlShenpiWorkBook);
            releaseObject(xlShenpiApp);
        }


        //根据stack,判断这一天的状态,  tmp_stack 里面存放的是一天的打卡数据
        //出勤：/   加班：加   旷工：旷   迟到：迟   早退：退   事假：事   病假：病   产假：产   婚假：婚   丧假：丧   出差：差   外出：外  调休：调  未打卡：未（备注栏填写次数）
        private void JudgeOneDayStatus(ref Stack<DateTime> tmp_stack, ref DateTime _today, ref OneUserDayStatus one_status, ref string user_name)
        {
            //先判断这一天是周几
            switch (tmp_stack.Count) 
            {
                case 0:             //先判断栈里面有几个元素,如果为0,表示这一天没有打卡记录
                    if (_today.DayOfWeek == DayOfWeek.Saturday || _today.DayOfWeek == DayOfWeek.Sunday) //说明是周六或者周天
                        one_status.SetValue(_today.Day, user_name, "", "");
                    else
                        one_status.SetValue(_today.Day, user_name, "旷", "8H");
                    break;
                case 1:  //这一天只有一条数据,一律认为旷工0.5天
                        one_status.SetValue(tmp_stack.Peek().Day, user_name, "旷", "4H");
                    break;
                default : //这一天有两条或者两条以上记录是最正常的状态
                    DateTime last_record = tmp_stack.Peek();//栈顶,最后下班的时间
                    DateTime first_record  = tmp_stack.ElementAt(tmp_stack.Count - 1);//栈底,最早上班打卡时间
                    /*1. 正常状态   在10点之前打卡,并且   下班打卡时间 - 上班打卡时间  >= 9个小时,例如   9:23上班打卡, 下班需要在 18:23以及以后打卡
                      2. 上班迟到   在10点之后打卡,并且    下班打卡时间 - 上班打卡时间 >= 9个小时, 例如  10:05上班打卡, 下班在19:10打卡, 迟到5分钟 
                      3. 下班早退(也算迟)   在10点之前打卡,并且    下班时间 - 上班打卡时间  < 9个小时, 例如   9:30上班打卡, 下班时间应为18:30,但是如果 是18:25下班,早退5分钟
                      4.  上班迟到/下班早退,上班打卡时间在10点之后打卡的,  一定会计入迟到时间,   下班在19:00及其以后打卡的不计入迟到时间, 下班时间在 19:00之前打卡的 都计入迟到时间   */
                    int N = 0;
                    if ((first_record.Hour < 10 || first_record.Hour ==10 && first_record.Minute == 0 )  &&  (last_record - first_record).TotalMinutes >= 9 * 60 )  //10点之前打卡,并且下班时间足够
                        one_status.SetValue(tmp_stack.Peek().Day, user_name, "/", "");
                    else if ((first_record.Hour > 10 || first_record.Hour == 10 && first_record.Minute > 0) && (last_record - first_record).TotalMinutes >= 9 * 60 )  //10点之后打卡,下班时间足够, 迟到
                        one_status.SetValue(tmp_stack.Peek().Day, user_name, "迟", ((first_record.Hour - 10)*60 + first_record.Minute).ToString() );
                    else if ((first_record.Hour < 10 || first_record.Hour == 10 && first_record.Minute == 0) && (last_record - first_record).TotalMinutes < 9 * 60)  // 上班时间正常,下班提前走了
                    {
                        N = (int)((first_record.AddHours(9) - last_record).TotalMinutes);
                        if( N == 0)
                            one_status.SetValue(tmp_stack.Peek().Day, user_name, "/", "");
                        else
                            one_status.SetValue(tmp_stack.Peek().Day, user_name, "迟", N.ToString());
                    }
                    else if ((first_record.Hour > 10 || first_record.Hour == 10 && first_record.Minute > 0) && (last_record.Hour > 19 || (last_record.Hour == 19 && last_record.Minute >= 0)))
                        one_status.SetValue(tmp_stack.Peek().Day, user_name, "迟", ((first_record.Hour - 10) * 60 + first_record.Minute).ToString());
                    else if ((first_record.Hour > 10 || first_record.Hour == 10 && first_record.Minute > 0) && last_record.Hour < 19)
                        one_status.SetValue(tmp_stack.Peek().Day, user_name, "迟", ((first_record.Hour - 10) * 60 + first_record.Minute + (19 - last_record.Hour) * 60 - last_record.Minute).ToString());
                    break;
            } 
        }

        ///////////////////////////////////////////excel表格操作函数/////////////////////
        //获取指定单元格的内容
        public string GetCellContent(int row, int column, ref Worksheet myExcel)
        {
            string tem = "";
            Excel.Range c1 = myExcel.Cells[row, column];
            Excel.Range c2 = myExcel.Cells[row, column];
            Excel.Range range = (Excel.Range)myExcel.get_Range(c1, c2);
            tem = Convert.ToString(range.Value);
            return tem;
        }

        public void SetCellColor(int row, int column, ref Worksheet myExcel)
        {
            Excel.Range c1 = myExcel.Cells[row, column];
            Excel.Range c2 = myExcel.Cells[row, column];
            Excel.Range range = (Excel.Range)myExcel.get_Range(c1, c2);
            range.Interior.ColorIndex = 36;
        }

        public void GetFirstValidUserName(ref Worksheet myExcel, out string name, out int valid_idx, int iRowsCount)
        {
            int valid_index = 0; //有效数据的起始行
            string tmp_name = "";
            //找到第一个考勤号码大于等于10000的那一行,然后获取D列的日期
            for (int i = 2; i <= iRowsCount; i++)  //从第二行开始,第一行是标题
            {
                string colum_C_data = ""; //c列的数据
                colum_C_data = GetCellContent(i, 3, ref myExcel);
                if (Int32.Parse(colum_C_data) >= 10000)
                {
                    valid_index = i;
                    tmp_name = GetCellContent(i, 2, ref myExcel);
                    break;
                }
            }
            name = tmp_name;
            valid_idx = valid_index;
        }

        //获取考勤数据中最早和最晚的时间,获取三个人的考勤记录,得到最早和最晚
        public void GetEarlistLatestTime(ref Worksheet myExcel, out string ealist, out string latest, int valid_idx, int iRowsCount, string first_name) 
        {
            string tmp_string = first_name;
            int count = 0;
            //读取3个不同用户的数据
            List<string> list = new List<string>();
            // 从有效行开始读取
            for (int i = valid_idx; i < iRowsCount; i++) 
            {
                list.Add(GetCellContent(i, 4, ref myExcel));
                string name = GetCellContent(i, 2, ref myExcel);
                if (name != tmp_string)
                {
                    count++;
                    tmp_string = name;
                    if (count == 3)
                        break;
                }
            }
            list.Sort();
            ealist = list[0];
            latest = list[list.Count - 1];

        }

        //添加标注
        public bool AddComent(object coment, int row, int column, ref Worksheet myExcel)
        {
            try
            {
                //Microsoft.Office.Interop.Excel.Range range = range[rowIndex, columnIndex] as Microsoft.Office.Interop.Excel.Range
                Excel.Range c1 = myExcel.Cells[row, column];
                Excel.Range c2 = myExcel.Cells[row, column];
                Excel.Range range = (Excel.Range)myExcel.get_Range(c1, c2);

                if (range.Comment != null)
                {
                    range.Comment.Delete();
                }
                range.AddComment(coment);
                return true;
            }
            catch
            {
                return false;
            }
        }

        //释放对象
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        //跨线程方位ui界面空间
        delegate void SetDateTimeCallback(DateTime dt);

        private void SetDateTimePickerValue(DateTime dt)
        {
            if (this.dateTimePicker_start.InvokeRequired)
            {
                SetDateTimeCallback stc = new SetDateTimeCallback(SetDateTimePickerValue);
                this.Invoke(stc, new object[] { dt});
            }
            else 
            {
                this.dateTimePicker_start.Value = dt;
            }
        }

        private void UpdateDateTimePickerStart(DateTime dt)
        {
            if (this.dateTimePicker_start.InvokeRequired)
            {
                // 当一个控件的InvokeRequired属性值为真时，说明有一个创建它以外的线程想访问它
                Action<DateTime> actionDelegate = (x) => { this.dateTimePicker_start.Value = x; };
                // 或者
                // Action<string> actionDelegate = delegate(string txt) { this.label2.Text = txt; };
                this.dateTimePicker_start.Invoke(actionDelegate, dt);
            }
            else
            {
                this.dateTimePicker_start.Value = dt;
            }
        }

        private void UpdateDateTimePickerEnd(DateTime dt)
        {
            if (this.dateTimePicker_end.InvokeRequired)
            {
                // 当一个控件的InvokeRequired属性值为真时，说明有一个创建它以外的线程想访问它
                Action<DateTime> actionDelegate = (x) => { this.dateTimePicker_end.Value = x; };
                // 或者
                // Action<string> actionDelegate = delegate(string txt) { this.label2.Text = txt; };
                this.dateTimePicker_end.Invoke(actionDelegate, dt);
            }
            else
            {
                this.dateTimePicker_end.Value = dt;
            }
        }

        //在最终写入的表格中,寻找对应的那个人
        public void WriteToFinalExcel(ref List<OneUserDayStatus> _oneday_status, string _user_name)
        { 
            //从最终表的第五行开始寻找, 寻找E列的数据 (即第5列,姓名),一旦找到对应的行数,从F列开始写入
            for (int i = 5; i <= need_write_sheet.UsedRange.Rows.Count; i++) {
                if (_user_name == GetCellContent(i, 5, ref need_write_sheet)) 
                {
                    int weidaka_times = 0;
                    //从list的0元素开始遍历,例如 list中的数据是 从 7号开始,那么还需要对单元格的列进行循环
                    for (int j = 0; j < _oneday_status.Count ; j++) 
                    { 
                        //依次获取day_of_month的数据
                        int column = _oneday_status.ElementAt(j).day_of_month + 5;
                        if (GetCellContent(i, column, ref need_write_sheet) == null)//不是放假的数据才写入
                        {
                            string daka_staus = _oneday_status.ElementAt(j).status;
                            if (!week_end_list.Contains(column)) //不在week_end_list中的才写入到cell中
                            {
                                need_write_sheet.Cells[i, column] = daka_staus;
                                if (_oneday_status.ElementAt(j).extra_data != "")
                                    AddComent(_oneday_status.ElementAt(j).extra_data, i, column, ref need_write_sheet);
                                if (daka_staus != "/") //不是 /   都加粗
                                    need_write_sheet.Cells[i, column].Font.Bold = true;
                                //如果打卡状态里面包含  未  字,标记为黄色
                                if (daka_staus.Contains("未"))
                                {
                                    need_write_sheet.Cells[i, column].Interior.Color = Color.Yellow;
                                    weidaka_times++;
                                }
                            }
                        }
                    }//内存for循环
                    break;
                }//if判断
            }//for外层
        }

        private void btn_generate_final_Click(object sender, EventArgs e)
        {
            object misValue = System.Reflection.Missing.Value;
            need_write_book.Close(true, misValue, misValue);
            need_write_app.Quit(); //这一句很重要,否则excel对象不能从内存中退出
            //释放资源
            releaseObject(need_write_sheet);
            releaseObject(need_write_book);
            releaseObject(need_write_app);
        }

        //tmp_s中存放的是符合日期要求的,一个用户的所有打开记录, _stack_oneday存放的是一个用户一天的所有打开记录
        private void GetOneDayStackData(ref Queue<DateTime> tmp_q, string _user_name)
        {
            //从时间控件start到end进行循环
            //for (DateTime dt = this.dateTimePicker_start.Value.Date; dt <= this.dateTimePicker_end.Value.Date; dt = dt.AddDays(1))
           // {
            Stack<DateTime> _stack_oneday = new Stack<DateTime>();

            for (DateTime dt = this.dateTimePicker_start.Value.Date; dt <= this.dateTimePicker_end.Value.Date; dt = dt.AddDays(1))
             {
                     foreach (DateTime one_day_datetime in tmp_q)
                     {
                         if (one_day_datetime.Date.Day == dt.Date.Day)
                         {
                             _stack_oneday.Push(one_day_datetime);
                         }

                     }
                 OneUserDayStatus one_status_obj = new OneUserDayStatus();
                 JudgeOneDayStatus(ref _stack_oneday, ref dt,ref one_status_obj, ref _user_name);
                 //获取完这一天的数据之后,把栈清空
                 _stack_oneday.Clear();
                 oneday_list.Add(one_status_obj);
                 Console.WriteLine("日期序号: {0} 员工:{1} 今日状态: {2}  标注数据: {3}", one_status_obj.day_of_month, one_status_obj.user_name,
                     one_status_obj.status, one_status_obj.extra_data);

             }


        }

        private void btn_gen_acc_chidao_Click(object sender, EventArgs e)
        {
            //启动一个子线程来处理考勤数据
            ThreadStart Ts = new ThreadStart(ProcessAccChidao);
            Thread Thread_acc = new Thread(Ts);
            Thread_acc.Start();
            Thread_acc.IsBackground = true;
        }

        //从考勤矩阵中获取迟到次数和累积迟到时长
        struct acc_chidao
        {
            public string name;
            public int chidao_times;//迟到次数
            public int acc_chidao_period;//累积迟到时长
        };

        private void ProcessAccChidao() 
        {
            List<acc_chidao> acc_list = new List<acc_chidao>();
           for (int i = 5; i <= 228; i++)
           {
               acc_chidao one_acc_chidao = new acc_chidao();
               one_acc_chidao.chidao_times = 0;
               one_acc_chidao.acc_chidao_period = 0;

               string user_name = GetCellContent(i,5,ref need_write_sheet);
               for (int j = 6; j <= 33; j++) 
               {
                   one_acc_chidao.name = GetCellContent(i, 5, ref need_write_sheet);
                   string status = GetCellContent(i, j, ref need_write_sheet);
                   
                   if (status =="迟")
                   {
                       one_acc_chidao.chidao_times++;
                       //Excel.Range c1 = need_write_sheet.Cells[i, j];
                       //Excel.Range c2 = need_write_sheet.Cells[i, j];
                       //Excel.Range range = (Excel.Range)need_write_sheet.get_Range(c1, c2);
                       Range range = (Range)need_write_sheet.Cells[i, j];
                       one_acc_chidao.acc_chidao_period += int.Parse(range.Comment.Text().ToString().Trim());
                   }
                   else if (status != null && status != "迟" && status.Contains("迟"))
                   {
                       one_acc_chidao.chidao_times++;
                       Console.WriteLine("i : {0}", i);
                       Range range = (Range)need_write_sheet.Cells[i, j];
                       string[] split = range.Comment.Text().ToString().Split(new Char[] { '/' });
                       if (split[0].Contains("分钟"))
                           split[0] = split[0].Substring(0, split[0].Length - 2);
                       one_acc_chidao.acc_chidao_period += int.Parse(split[0]);
                       
                   }
               }
               acc_list.Add(one_acc_chidao);
               
           }
            //读取
           Excel.Application xlApp;
           Excel.Workbook xlWorkBook;
           Excel.Worksheet xlWorkSheet;
           object misValue = System.Reflection.Missing.Value;

           xlApp = new Excel.Application();
           xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\st\Desktop\a.xls", 0, true,
               5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
           xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
           //把数据写入excel表
           for(int k=4;k<=227;k++)
           {
               string name = GetCellContent(k, 4, ref xlWorkSheet);
               for (int m = 0; m < acc_list.Count; m++)
               {
                   if (acc_list.ElementAt(m).name == name)
                   {
                       xlWorkSheet.Cells[k, 7] = acc_list.ElementAt(m).chidao_times;
                       xlWorkSheet.Cells[k, 8] = acc_list.ElementAt(m).acc_chidao_period;
                   }
               }
               
           }

           xlWorkBook.Close(true, misValue, misValue);
           xlApp.Quit(); //这一句很重要,否则excel对象不能从内存中退出
           //释放资源
           releaseObject(xlWorkSheet);
           releaseObject(xlWorkBook);
           releaseObject(xlApp);


        }

        
        

    }
}
