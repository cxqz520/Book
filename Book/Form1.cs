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


namespace Book
{
    public partial class Form1 : Form
    {
        public static string selectedType,quarter,selectedForm,selectedForm2=null,ACCode= "GB2312";//"GB2312" "UTF-8"
        public static int selectedYear, selectedMonth=01;
        public static String [] typeList = { "住房", "餐饮","衣服", "妆品", "车费", "日用", "电子" };
        public static String[] quarterList = {  "衣服", "妆品", "车费", "日用" };
        public static String[] incomeTypeList = { "工资", "红包", "投资", "代购"};
        public DataTable fileMonth;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            BackUp();

            String[] sheetList = { "收支总表", "重要支出", "预算管理", "预算对比" };//下拉框选项
            String[] sheetList3 = { "衣服", "妆品", "日用", "书籍电子", "额外" };//下拉框选项

            int i;

            //************初始化界面**********************//
            //获取列表
            for (i = 0; i < sheetList.Length; i++) comboBox1.Items.Add(sheetList[i]);//遍历选择框1 combox1
            comboBox1.SelectedIndex = 0;//初始值
            for (i = 0; i < sheetList3.Length; i++) comboBox3.Items.Add(sheetList3[i]);//遍历选择框3 combox3
            comboBox3.Items.Add("All Selection");
            comboBox3.SelectedIndex = 0;//初始值
            initMonthSheetByYear();
        }


        ///***************初始化界面方法**********************//
        public void loadGridView()
        {
            String sql, tmpStr;
            switch (this.comboBox1.SelectedIndex)
            {
                case 0:
                    this.Width = 1310;
                    dataGridView2.Visible = true;
                    lableIncome.Visible = true;
                    comboBox3.Visible = false;
                    button2.Visible = true;
                    button3.Visible = false;
                    LoadYearSelector();//录入年份选择框 combox2
                    //this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    sql = "Date Like '" + selectedYear + "%' ";
                    LoadMonthSheet(sql);
                    addSumLine(1);
                    dataGridView1.ReadOnly = false;
                    dataGridView1.Columns[0].ReadOnly = true;
                    dataGridView2.Columns[0].ReadOnly = true;
                    labeloutCome.Text = "总支出： "+getSumFromGridColounm(dataGridView1,dataGridView1.Columns.Count-1);
                    addSumRow("Sum");
                    addSumRow("Budget");
                    break;
                case 1://重要支出
                    this.Width = 970;
                    dataGridView2.Visible = false;
                    lableIncome.Visible = false;
                    comboBox3.Visible = true;
                    button2.Visible = true;
                    button3.Visible = true;
                    LoadYearSelector();//录入年份选择框 combox2
                    dataGridView1.ReadOnly = false;
                    if (comboBox2.SelectedIndex == -1)
                    {
                        sql = "品类 Like '" + selectedType + "%'";
                    }
                    else if(comboBox3.SelectedItem.ToString() == "All Selection")
                    {
                        sql = "Date Like '" + selectedYear + "%'";
                    }
                    else{ sql = "Date Like '" + selectedYear + "%' and 品类 Like '" + selectedType + "%'"; }
                    LoadImportantSheet("detail", sql);
                    labeloutCome.Text = "总支出： " + getSumFromGridColounm(dataGridView1, 2);
                    break;
                case 2://"预算管理"
                    ClearDataGrid();
                    this.Width = 970;
                    dataGridView2.Visible = false;
                    lableIncome.Visible = false;
                    comboBox3.Visible = true;
                    comboBox2.DataSource = null;
                    button2.Visible = true;
                    button3.Visible = true;
                    dataGridView1.ReadOnly = false;
                    initMonthSheetByYear();
                    LoadYearSelector();//录入年份选择框 combox2
                    if (comboBox3.SelectedItem.ToString() == "All Selection")
                    {
                        sql = "Date Like '" + selectedYear + "%'";
                    }else sql = "Date Like '" + selectedYear + "%' and 品类 Like '" + selectedType + "%'";
                    LoadImportantSheet("budget", sql);

                    break;
                case 3://"预算对比"
                    ClearDataGrid();
                    this.Width = 970;
                    dataGridView2.Visible = false;
                    lableIncome.Visible = false;
                    comboBox3.Visible = false;
                    button2.Visible = false;
                    LoadQuarterSelector();//录入年份选择框 combox2
                    tmpStr = comboBox2.SelectedItem.ToString();
                    selectedYear = Convert.ToInt32(tmpStr.Substring(0, 4));
                    //quarter = tmpStr.Substring(tmpStr.Length - 2);
                    //selectedYear = ;
                    dataGridView1.ReadOnly = true;
                    if (tmpStr.Length > 4)
                    {
                        selectedMonth = Convert.ToInt32(tmpStr.Substring(5));
                        sql = "Date ='" + selectedYear.ToString() + "." + GetMonth2Dig(selectedMonth).ToString() + "'";
                        //"Date >= 2019.1 and Date<2019.4 " ??2019.11/2019.12?
                    }
                    else
                    {
                        selectedMonth = 0;
                        sql = "";
                    }
                    LoadCompareSheet(sql);
                    addSumRow("Compare");
                    break;
            }
        }

        public void initMonthSheetByYear()
        {
            DataTable table = new DataTable("month");
            table.Columns.Add("Date");
            for (int i = 0; i < typeList.Length; i++)
            { //遍历添加列
                table.Columns.Add(typeList[i]);//支出
            }
            for (int i = 0; i < incomeTypeList.Length; i++)
            { //遍历添加列
                table.Columns.Add(incomeTypeList[i]);//收入
            }

            for (int i=1;i<=12;i++)
            {
                table.Rows.Add(new[] { selectedYear + "." + GetMonth2Dig(i), "0","0","0","0","0","0","0","0","0","0","0"});
            }

            //table = null;
            DataTable budgetTmp = new DataTable();
            budgetTmp.Columns.Add("Date");//时间
            budgetTmp.Columns.Add("品类");
            budgetTmp.Columns.Add("分类");
            budgetTmp.Columns.Add("金额");
            budgetTmp.Columns.Add("备注");
            for (int i=0; i<typeList.Length; i++)
            {
                budgetTmp.Rows.Add(new[] { selectedYear+"", typeList[i], "month", "0", "" });
                if (quarterList.Contains(typeList[i]))
                {
                    budgetTmp.Rows.Add(new[] { selectedYear + "", typeList[i], "quarter", "0", "" });
                }
            }


            SaveTableByLine(RemoveRepeatLine(budgetTmp, "budget", 3), "budget");
            SaveTableByLine(RemoveRepeatLine(table, "month", 1), "month");
            budgetTmp.Dispose();
            table.Dispose();
        }
        //*************************方法/Function*****************************//
        public void LoadYearSelector()
        {//录入年份
            int currentYear = DateTime.Now.Year;
            if (comboBox2.Items.Contains(currentYear.ToString() + "-01"))
            {
                comboBox2.Items.Clear();
            }

            if (comboBox2.Items.Contains(currentYear.ToString())) {
            }
            else
            {
                for (int i = 2018; i <= currentYear; i++)
                {
                    comboBox2.Items.Add(Convert.ToString(i));
                }
                comboBox2.SelectedItem = currentYear+"";
                selectedYear = currentYear;
            }
        }
        public void LoadQuarterSelector()
        {//录入年份
            int currentYear = DateTime.Now.Year;
            int currentMonth = DateTime.Now.Month;
            if (comboBox2.Items.Contains(currentYear.ToString() + "-01"))
            { } else
            {
                for (int i = currentYear; i >= 2018; i--)
                {
                    for (int j = 1; j <= 12; j++)
                    {
                        comboBox2.Items.Add(Convert.ToString(i) + "-"+ GetMonth2Dig(j));
                    }
                }
                comboBox2.SelectedItem = currentYear+"-"+ currentMonth;
                selectedYear = currentYear;
                selectedMonth = Convert.ToInt32(GetMonth2Dig(currentMonth));
            }
        }
        public void addSumLine(int index)
        {
            double tmp = 0;
            String tmpStr;
            DataGridViewTextBoxColumn cl = new DataGridViewTextBoxColumn();
            cl.Name = "total";
            cl.HeaderText = "total";
            this.dataGridView1.Columns.Add(cl);
            //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                tmp = 0;
                //MessageBox.Show("sss" + this.dataGridView1.ColumnCount);
                for (int j = 1; j < (this.dataGridView1.ColumnCount - 1); j++)//第一格年份 最后一格新加的
                {
                    tmpStr = this.dataGridView1.Rows[i].Cells[j].Value.ToString();

                    tmp += double.Parse(tmpStr);
                }
                //MessageBox.Show("sss" + tmp);
                this.dataGridView1.Rows[i].Cells["total"].Value = tmp;
            }
        }
        public void addSumRow(string type)
        {

            switch (type)
            {
                case "Sum":
                    AddNewRow(dataGridView1);
                    AddNewRow(dataGridView2);
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = "Sum for Type";
                    dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[0].Value = "Sum for Type";
                    //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
                    for (int i = 1; i < this.dataGridView1.Columns.Count; i++)
                    {
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value = getSumFromGridColounm(dataGridView1, i);
                    }
                    for (int i = 1; i < this.dataGridView2.Columns.Count; i++)
                    {
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[i].Value = getSumFromGridColounm(dataGridView2, i);
                    }

                    break;
                case "Budget":
                    AddNewRow(dataGridView1);
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = "Budget Sum";
                    //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
                    for (int i = 1; i < this.dataGridView1.Columns.Count-1; i++)
                    {
                        double tmp = getBudgetToMonth((selectedYear == DateTime.Now.Year ? DateTime.Now.Month : 12), dataGridView1.Columns[i].Name) - Convert.ToDouble(dataGridView1.Rows[13].Cells[i].Value);
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value = Math.Round(tmp,2);
                    }
                    break;
                case "Compare":
                    //AddNewRow(dataGridView1);
                    double tmpSum=0;
                    dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = "Budget Sum";
                    //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
                    for (int i = 2; i < 6; i++)
                    {
                        for(int j = 0; j < dataGridView1.Rows.Count; j++)
                        {
                            tmpSum += Convert.ToDouble( dataGridView1.Rows[j].Cells[i].Value);
                        }
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[i].Value = tmpSum;
                        tmpSum = 0;
                    }
                    break;
            }
            
        }

        public void BackUp()
        {
            String date = "BackUp_Data_"+DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss");
            //MessageBox.Show("this is path" + Application.StartupPath);
            String aimPath = Application.StartupPath + "\\" + date;
            System.IO.Directory.CreateDirectory(aimPath);//创建文件夹
            String searchPath = Application.StartupPath + "\\" + "Data";
            DirectoryInfo dir = new DirectoryInfo(searchPath);
            FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();
            foreach (FileSystemInfo i in fileinfo)
            {
                File.Copy(i.FullName, aimPath + "\\" + i.Name, true);
            }

            }
        public double getBudgetToMonth(int month, string type)
        {
            double budgetTotal = month * GetBudget(selectedYear, type , "month")+ (int)(month/3+1)* GetBudget(selectedYear, type, "quarter");
            
            return budgetTotal;
        }
        public double getSumByTime(int year,int month, string type)
        {
            double sum = 0;
            string sql = "Date<='" + year + "." + GetMonth2Dig(month) + "' and Date>='" + year + ".00'";
            if (fileMonth.Select(sql).Length > 0) {
                DataTable temp = new DataTable();
                temp = fileMonth.Select(sql).CopyToDataTable();
                for (int i=0;i<temp.Rows.Count;i++)
                {
                    sum = sum+ Convert.ToDouble(temp.Rows[i][type]);

                }
                temp = null;
            } 
            return sum;
        }
        public double getSumFromGridColounm(DataGridView dataGridView, int columnIndex)
        {
            double sum = 0;
            //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                //MessageBox.Show("sss" + this.dataGridView1.ColumnCount);
                sum += Convert.ToDouble(dataGridView.Rows[i].Cells[columnIndex].Value.ToString());
            }

            return sum;
        }
        public void LoadDataGrid(DataTable table, int fromLine, int toLine, DataGridView dataGrid)
        {// 将指定table的内容输出到 dataGridView （按列增加）
            int i;
            //将表头输出
            //MessageBox.Show("This is table row count" + table.Rows.Count);
            for (i = fromLine; i <= toLine && i < table.Columns.Count; i++)
            {
                DataGridViewTextBoxColumn cl = new DataGridViewTextBoxColumn();
                cl.Name = table.Columns[i].ColumnName;
                cl.HeaderText = table.Columns[i].ColumnName;

                dataGrid.Columns.Add(cl);
            }
            //遍历输出行
            for (i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows.Count > dataGrid.Rows.Count)
                    dataGrid.Rows.Add();


                for (int j = fromLine; j <= toLine && j < table.Columns.Count; j++)
                {
                    if (j-fromLine < dataGrid.ColumnCount)
                        
                        dataGrid.Rows[i].Cells[j - fromLine].Value = table.Rows[i][table.Columns[j].ColumnName].ToString();//便利格子
                }
            }
        }
        public void LoadDataGrid(DataTable table, int fromLine, int toLine, DataGridView dataGrid, int fromDGLine)
        {// 将指定table的内容输出到 dataGridView （按列增加）
            int i;
            //将表头输出
            //MessageBox.Show("This is table row count" + table.Rows.Count);
            for (i = fromLine; i <= toLine && i < table.Columns.Count; i++)
            {
                DataGridViewTextBoxColumn cl = new DataGridViewTextBoxColumn();
                cl.Name = table.Columns[i].ColumnName;
                cl.HeaderText = table.Columns[i].ColumnName;

                dataGrid.Columns.Add(cl);
            }
            //遍历输出行
            for (i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows.Count > dataGrid.Rows.Count)
                    dataGrid.Rows.Add();


                for (int j = fromLine; j <= toLine && j < table.Columns.Count; j++)
                {
                    if (j - fromLine < dataGrid.ColumnCount)

                        dataGrid.Rows[i].Cells[j+ fromDGLine - fromLine].Value = table.Rows[i][table.Columns[j].ColumnName].ToString();//便利格子
                }
            }
        }
        public void InsertGridByLine(DataTable table, DataGridView dataGrid)//将data table导入grid下方
        {//将指定table的内容输出到 dataGridView 按行增加
            //String title;
            if (dataGrid.Columns.Count < table.Columns.Count)
            {
                if (dataGrid.Columns.Count < table.Columns.Count) { 
                    for (int i = dataGrid.Columns.Count; i < table.Columns.Count; i++)
                    {
                        DataGridViewTextBoxColumn cl = new DataGridViewTextBoxColumn();
                        cl.Name = table.Columns[i].ColumnName;//table 第i列的列名
                        cl.HeaderText = table.Columns[i].ColumnName;
                        dataGrid.Columns.Add(cl);
                    }
                }
            }
            int rowsCount = dataGrid.Rows.Count;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                dataGrid.Rows.Add();
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    if (j < dataGrid.ColumnCount)
                        dataGrid.Rows[rowsCount+i-1].Cells[j].Value = table.Rows[i][table.Columns[j].ColumnName].ToString();//便利格子
                }
            }
        }
        public void ClearDataGrid()
        {//清空 dataGridView
            for (int i = this.dataGridView1.Columns.Count - 1; i >= 0; i--)
                this.dataGridView1.Columns.RemoveAt(i);
            for (int i = this.dataGridView2.Columns.Count - 1; i >= 0; i--)
                this.dataGridView2.Columns.RemoveAt(i);
        }
        public string[] SplitStr(String str)

        {
            string[] list = str.Split(',');
            return list;
        }
        public List<String> GetStringArray(String fileName)
        {//读取文件，返回文件内容
            var file = File.Open(fileName, FileMode.Open);
            List<string> txt = new List<string>();
            using (var stream = new StreamReader(file, Encoding.GetEncoding(ACCode)))
            {
                while (!stream.EndOfStream)
                {
                    txt.Add(stream.ReadLine());
                }
            }
            file.Close();
            return txt;
        }

        public string GetMonth2Dig(int month)
        {
            String str = "0" + month;
            str = str.Substring(str.Length-2,2);
            return str;
        }

        /*****************将文件存入Datatable*****************/
        public DataTable LoadFileToDT(String fileName)
        {//将文件存入Datatable
            int i, j;
            fileName = fileName + ".txt";
            DataTable table = new DataTable(fileName);
            fileName = @"Data\" + fileName;
            var tmpList = GetStringArray(fileName);
            var titleList = SplitStr(tmpList[0]);//获取第一列： 列表名

            for (i = 0; i < titleList.Length; i++) { //遍历添加列
                table.Columns.Add(titleList[i]);
            }

            //MessageBox.Show("这是数据的列"+tmpList.Count);
            for (i = 1; i < tmpList.Count; i++)//遍历行tmpList.Count
            {
                var tempLine = SplitStr(tmpList[i]);
                var row = table.NewRow();
                for (j = 0; j < titleList.Length && j<tempLine.Length; j++)//遍历列 titleList.Length
                {
                    /*if (table.Columns[j].ColumnName == "Date")
                    {
                        DateTime dateTime = DateTime.ParseExact(tempLine[j],"yyyy.m",null);
                        row[titleList[j]] = tempLine[j];
                    }
                    else
                    {
                        row[titleList[j]] = tempLine[j];
                    }*/
                    row[titleList[j]] = tempLine[j];
                }
                table.Rows.Add(row);
                row = null;
                tempLine = null;
            }
            tmpList = null;
            
            return table;
            ///可添加 容器的释放

        }

        public void LoadImportantSheet(String fileName, String sql)//Important
        {//加载页面 单表
            ClearDataGrid();
            DataTable detailSheet = LoadFileToDT(fileName);//将表导出
            DataRow[] Drow = detailSheet.Select(sql);
            if (Drow.Length == 0)
            {
                MessageBox.Show("Data not found");
                LoadDataGrid(detailSheet, 0, 5, dataGridView1);//将文件导入
            }
            else
            {
                DataTable Tmp = Drow.CopyToDataTable();
                LoadDataGrid(Tmp, 0, 5, dataGridView1);//将文件导入
                Tmp.Dispose();
            }

            detailSheet.Dispose();
            Drow = null;
        }
        public void LoadMonthSheet(String sql)
        {//加载页面 单表双视图
            ClearDataGrid();
            fileMonth = LoadFileToDT("month");
            //DataTable monthSheet = LoadFileToDT("month");//将表导出
            DataRow[] Drow = fileMonth.Select(sql);
            if (Drow.Length == 0)
            {
                MessageBox.Show("Data not found");
                LoadDataGrid(fileMonth, 0, 7, dataGridView1);//将文件导入
                LoadDataGrid(fileMonth, 0, 0, dataGridView2);//将文件导入
                LoadDataGrid(fileMonth, 8, 12, dataGridView2);//将文件导入
            }
            else
            {
                DataTable Tmp = Drow.CopyToDataTable();
                LoadDataGrid(Tmp, 0, 7, dataGridView1);//将文件导入
                LoadDataGrid(Tmp, 0, 0, dataGridView2);//将文件导入
                LoadDataGrid(Tmp, 8, 12, dataGridView2,1);//将文件导入
                Tmp.Dispose();
            }

        }
        public void LoadCompareSheet(String sql)// Date=2019.1
        {//加载页面 双表
            ClearDataGrid();
            String sql2,tmp;
            double dbudget, dextra, dactual;
            DataTable budgetSheet = LoadFileToDT("budget");//将表导出
            DataTable monthSheet = LoadFileToDT("month");//将表导出
            //DataTable budgetYear = GetBudgetTable(budgetSheet);// budgetYear 此为年度 每月/每季度 预算
            String[] title = { "Month", "Type", "Budget", "Extra", "Actual","Rest" };
            DataTable dtAll = new DataTable("Compare");//最后的表
            //DataTable typeTable = getTypeList(budgetSheet);
            for (int i = 0; i < title.Length; i++)
            {
                dtAll.Columns.Add(title[i]);//添加Compare表头
            }

            for (int i = 0; i < typeList.Length; i++)
            {
                var row = dtAll.NewRow();
                row["Month"] = selectedYear + "." + GetMonth2Dig(selectedMonth);
                row["Type"] = typeList[i];
                sql2 = "type='" + typeList[i] + "'";
                tmp = Convert.ToString(GetBudget(selectedYear, typeList[i], "month"));//GetExpectCell(budgetYear, sql2, "month");
                dbudget = Convert.ToDouble(tmp == "" ? "0" : tmp);//去空值
                tmp = Convert.ToString(GetBudget(selectedYear, typeList[i], "quarter"));// GetExpectCell(budgetYear, sql2, "quarter");
                dextra = Convert.ToDouble(tmp == "" ? "0" : tmp);//去空值
                sql2 = "Date='" + selectedYear + "." + selectedMonth + "' and 品类='" + typeList[i] + "'";
                tmp = GetExpectCell(monthSheet, sql, typeList[i]);
                dactual = Convert.ToDouble(tmp == "" ? "0" : tmp);//实际支出
                row["Budget"] = dbudget;
                dextra = Math.Round((getBudgetToMonth(selectedMonth - 1, typeList[i]) - getSumByTime(selectedYear, selectedMonth, typeList[i])), 2);
                row["Extra"] = dextra;
                row["Actual"] = dactual;
                row["Rest"] = dbudget+dextra-dactual;
                dtAll.Rows.Add(row);
            }
            InsertGridByLine(dtAll,dataGridView1);

            //budgetYear.Dispose();
            monthSheet.Dispose();
            budgetSheet.Dispose();
        }
        public String GetExpectCell(DataTable table, String sql, string coloumName)
        {//

            DataRow[] tmp = table.Select(sql);
            String result = "";
            if (tmp.Length>0)
            {
                DataTable tmpTable = tmp.CopyToDataTable();
                result = tmpTable.Rows[0][coloumName].ToString();
                tmpTable.Dispose();
            }

            tmp = null;
            return result;
        }



        public double GetBudget(int Year, string type , string range)//type=quarter/month
        {
            DataTable budgetSheet = LoadFileToDT("budget");
            string sql = "Date='" + Year + "' and 品类='" + type + "'" + " and 分类 = '"+range+"'";
            double amount = 0;
            if (budgetSheet.Select(sql).Length > 0)
            {
                DataTable temp = new DataTable();
                temp = budgetSheet.Select(sql).CopyToDataTable();
                amount = Convert.ToDouble(temp.Rows[0]["金额"]);
                temp = null;
            }
            return amount;
        }

        public DataTable TransGridToTable(DataGridView dataGridView)
        {
            String name = dataGridView.Name;
            DataTable dataTable = new DataTable(name);

            for(int i = 0; i < dataGridView.Columns.Count; i++)
            {//添加列
                dataTable.Columns.Add(dataGridView.Columns[i].Name);
            }            
            //MessageBox.Show("这是数据的列"+tmpList.Count);
            for (int i = 0; i < dataGridView.Rows.Count; i++)//遍历行tmpList.Count
            {
                var row = dataTable.NewRow();
                for (int j = 0; j < dataGridView.Columns.Count; j++)//遍历列 titleList.Length
                {
                    row[j] = Convert.ToString(dataGridView.Rows[i].Cells[j].Value);
                }
                dataTable.Rows.Add(row);
            }

            return dataTable;
        }
        public DataTable Combine2GridToDataTable(DataGridView dataGridView1, DataGridView dataGridView2, int stopColumnGrid1, int startColumnGrid2, int stopLine)
        {//将两个datagridView 合成一个
            DataTable dataTable = new DataTable("Tmp");
            int tempColumnsCount1 = (dataGridView1.Columns.Count > stopColumnGrid1 ? stopColumnGrid1 : dataGridView1.Columns.Count);

            for(int i = 0; i < tempColumnsCount1; i++)
            {
                dataTable.Columns.Add(dataGridView1.Columns[i].Name);
            }
            for (int i = startColumnGrid2; i < dataGridView2.Columns.Count; i++)
            {
                dataTable.Columns.Add(dataGridView2.Columns[i].Name);
            }

            for (int i = 0; i < dataGridView1.Rows.Count && i<stopLine; i++)//遍历行tmpList.Count
            {
                var row = dataTable.NewRow();
                for (int j = 0; j < tempColumnsCount1; j++)//遍历列 titleList.Length
                {
                    if (Convert.ToString(dataGridView1.Rows[i].Cells[j].Value) != "")
                    {
                        row[j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                    else
                    {
                        row[j] = 0;
                    }
                }
                for (int j = startColumnGrid2; j < dataGridView2.Columns.Count; j++)//遍历列 titleList.Length
                {
                    string aimCellsValue = Convert.ToString(dataGridView2.Rows[i].Cells[j].Value);
                    row[tempColumnsCount1 + j - startColumnGrid2] = aimCellsValue;
                }
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }
        public DataTable CheckLineInFile(DataTable GridTable, String FileName, int index)//index: 从第几个逗号截止检查
        {//检查文件中是否有这一行
            int flag = 0,found=0;//found为1为找到，0为没找到

            DataTable FileTable = LoadFileToDT(FileName);
            int FilelineCount = FileTable.Rows.Count;
            for (int i = 0; i < GridTable.Rows.Count; i++)
            {//遍历想插入的表 GridTable的第i行
                for(int j=0; j< FilelineCount && found==0;j++)
                {//遍历源文件的行 FileTable的第j行
                    for(int k=0; k < index; k++)
                    {//遍历查询前三个单元格
                        if (FileTable.Rows[j][k].ToString() != GridTable.Rows[i][k].ToString())
                        {
                            flag = 1;//如果两个表目标行的前index格不相等，清空
                            break;
                        }
                    }
                    if (flag == 0)
                    {//如果前index格都相等
                        found = 1;
                        for (int k = 0; k < FileTable.Columns.Count; k++)
                        {//遍历将Grid表的该行插入File表
                            FileTable.Rows[j][k] = GridTable.Rows[i][k];
                        }
                        break;
                    }
                    flag = 0;
                }
                if (found == 0) { //如果没找到相同的行
                    DataRow dr = FileTable.NewRow();
                    for (int k = 0; k < FileTable.Columns.Count; k++)
                    {
                        dr[k] = GridTable.Rows[i][k];
                    }
                    FileTable.Rows.Add(dr);
                    dr = null;
                }
                found = 0;//重设found
            }


            return FileTable;
        }
        public void SaveTableByLine(DataTable dataTable, String FileName)
        {//按行将datatable存入文件fileName.txt
            //DataTable File =  LoadFileToDT(FileName);
            FileName = @"Data\" + FileName + ".txt";
            FileStream F = new FileStream(FileName, FileMode.Open, FileAccess.Write);
            F.SetLength(0);//清空文件
            StreamWriter wr = new StreamWriter(F,  Encoding.GetEncoding(ACCode));
            String tmpStr="";

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                tmpStr = tmpStr + dataTable.Columns[i].ColumnName.ToString()+",";
            }
            tmpStr = tmpStr.Substring(0,tmpStr.Length-1);
            wr.WriteLine(tmpStr, Encoding.GetEncoding(ACCode));
            tmpStr = "";

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for(int j = 0; j < dataTable.Columns.Count; j++)
                {
                    if (dataTable.Rows[i][j].ToString() != "") { 
                    tmpStr = tmpStr + dataTable.Rows[i][j].ToString()+",";
                    }
                    else
                    {
                        tmpStr = tmpStr + "0"+",";
                    }
                }
                tmpStr = tmpStr.Substring(0, tmpStr.Length - 1);
                wr.WriteLine(tmpStr, Encoding.GetEncoding(ACCode));
                tmpStr = "";
            }
            wr.Close();
            F.Close();

        }
        public DataTable RemoveRepeatLine(DataTable GridTable, String FileName, int index)
        {
            int flag = 0, found = 0;//found为1为找到，0为没找到
            //flag 0为前index个单元格一直相等
            DataTable FileTable = LoadFileToDT(FileName);
            int FilelineCount = FileTable.Rows.Count;

            for (int i = 0; i < GridTable.Rows.Count; i++)
            {//遍历想插入的表 GridTable的第i行
                for (int j = 0; j < FilelineCount && found == 0; j++)
                {//遍历源文件的行 FileTable的第j行
                    for (int k = 0; k < index; k++)
                    {//遍历查询前三个单元格
                        /*MessageBox.Show("Verify index " + k + " Grid table " + i + " line value is  " + GridTable.Rows[i][k].ToString() +
                            "  .  File Table " + j + " line value is" + FileTable.Rows[j][k].ToString());*/
                        if (FileTable.Rows[j][k].ToString() != GridTable.Rows[i][k].ToString())
                        {
                            //MessageBox.Show("not equal");
                               flag = 1;//如果两个表目标行的前index格不相等，清空
                            break;
                        }
                    }
                    if (flag == 0)
                    {//如果前index格都相等 即在表中找到相同行
                        //MessageBox.Show("Grid table " + j + "  line's value is equal File Table   " + i + " line's value");
                        found = 1;
                    }
                    flag = 0;
                }
                if (found == 0)//结束循环时仍没找到相同行， 则加入新行  
                {
                    DataRow dr = FileTable.NewRow();
                    for (int k = 0; k < FileTable.Columns.Count; k++)
                    {
                        dr[k] = GridTable.Rows[i][k].ToString();
                    }
                    FileTable.Rows.Add(dr);
                    dr = null;
                }
                flag = 0;//重设flag
                found = 0;//重设found
            }
            return FileTable;
        }

        public void AddNewRow(DataGridView dataGridView)
        {
            //int index = dataGridView.Rows.Add();
            int index = dataGridView.Rows.Count-1;
            dataGridView.Rows.Insert(index);

            for (int i=0;i<dataGridView.Columns.Count;i++)
            {
                dataGridView.Rows[index].Cells[i].Value = dataGridView.Rows[index+1].Cells[i].Value;
            }
            switch (this.comboBox1.SelectedIndex)
            {
                case 1:
                    dataGridView.Rows[index+1].Cells["Date"].Value = Convert.ToString(selectedYear)+ "." + GetMonth2Dig(DateTime.Now.Month);
                    dataGridView.Rows[index+1].Cells["品类"].Value = selectedType;
                    break;
                case 2:
                    dataGridView.Rows[index+1].Cells["Date"].Value = Convert.ToString(selectedYear);
                    dataGridView.Rows[index+1].Cells["品类"].Value = selectedType;
                    break;
            }
        }

        //**********************Function end***********************************//


        //************************事件************//
        private void Button2_Click(object sender, EventArgs e)
        {
            switch (this.comboBox1.SelectedIndex)
            {
                case 0:
                    SaveTableByLine(CheckLineInFile(Combine2GridToDataTable(dataGridView1, dataGridView2, 8, 1,12),"month",1),"month");
                    break;
                case 1:
                    SaveTableByLine(CheckLineInFile(TransGridToTable(dataGridView1), "detail", 4), "detail");//***需要加判断
                    break;
                case 2:
                    SaveTableByLine(CheckLineInFile(TransGridToTable(dataGridView1), "budget", 3), "budget");
                    break;
            }
            /*
             * 
             * for(int i = 0; i< )if()
            {

            }
            SaveGridByLine();*/
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            AddNewRow(dataGridView1);
        }

        private void DataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DialogResult dr=DialogResult.Cancel;
            switch (this.comboBox1.SelectedIndex)
            {
                case 1:
                    dr = MessageBox.Show("确定要删除该行吗？", "提示", MessageBoxButtons.OKCancel);
                    break;
                case 2:
                    dr = MessageBox.Show("确定要删除该行吗？", "提示", MessageBoxButtons.OKCancel);
                    break;
            }

            if (dr == DialogResult.OK)
            {
                dataGridView1.Rows.Remove(dataGridView1.Rows[e.RowIndex]);
            }
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            loadGridView();

        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(this.comboBox2.SelectedIndex != -1)
            {
                Form1.selectedYear = Convert.ToInt32(this.comboBox2.SelectedItem.ToString().Substring(0,4));
                loadGridView();
            }
        }

        private void ComboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if(comboBox3.SelectedIndex != -1)
            {
                Form1.selectedType = this.comboBox3.SelectedItem.ToString();
                loadGridView();
            }
            
        }

        private void DataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Button1_Click_1(object sender, EventArgs e)
        {

        }

    }
}
