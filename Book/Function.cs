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
    public class Function
    {
        public static string selectedType, quarter, selectedForm, selectedForm2 = null, ACCode = "GB2312";//"GB2312" "UTF-8"
        public Function()
        {


        }

        public void LoadYearSelector1(ComboBox comboBox)
        {//录入年份
            int currentYear = DateTime.Now.Year;
            if (comboBox.Items.Contains(currentYear.ToString() + "-01"))
            {
                comboBox.Items.Clear();
            }

            if (comboBox.Items.Contains(currentYear.ToString()))
            {
            }
            else
            {
                for (int i = 2018; i <= currentYear; i++)
                {
                    comboBox.Items.Add(Convert.ToString(i));
                }
                comboBox.SelectedItem = currentYear + "";
            }
            //return currentYear;
        }

        public void addSumLine(DataGridView dataGridView)
        {
            double tmp = 0;
            String tmpStr;
            DataGridViewTextBoxColumn cl = new DataGridViewTextBoxColumn();
            cl.Name = "total";
            cl.HeaderText = "total";
            dataGridView.Columns.Add(cl);
            //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                tmp = 0;
                //MessageBox.Show("sss" + this.dataGridView1.ColumnCount);
                for (int j = 1; j < (dataGridView.ColumnCount - 1); j++)//第一格年份 最后一格新加的
                {
                    tmpStr = dataGridView.Rows[i].Cells[j].Value.ToString();

                    tmp += double.Parse(tmpStr);
                }
                //MessageBox.Show("sss" + tmp);
                dataGridView.Rows[i].Cells["total"].Value = tmp;
            }
        }
        public void BackUp()
        {
            String date = "BackUp_" + DateTime.Now.ToString("yy.MM.dd.HH.mm.ss");
            //MessageBox.Show("this is path" + Application.StartupPath);
            String aimPath = Application.StartupPath + "\\BackUp\\" + date;
            System.IO.Directory.CreateDirectory(aimPath);//创建文件夹
            String searchPath = Application.StartupPath + "\\" + "Data";
            DirectoryInfo dir = new DirectoryInfo(searchPath);
            FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();
            foreach (FileSystemInfo i in fileinfo)
            {
                File.Copy(i.FullName, aimPath + "\\" + i.Name, true);
            }
        }

        public double getBudgetToMonth(int selectedYear, int month, string type)
        {
            double budgetTotal = month * GetBudget(selectedYear, type, "month") + (int)(month / 3 + 1) * GetBudget(selectedYear, type, "quarter");

            return budgetTotal;
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
        public void addSumRow(string type, DataGridView dataGridView, int selectedYear)
        {

            switch (type)
            {
                case "Sum":
                    AddNewRow(dataGridView);
                    dataGridView.Rows[dataGridView.Rows.Count - 1].Cells[0].Value = "Sum for Type";
                    //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
                    for (int i = 1; i < dataGridView.Columns.Count; i++)
                    {
                        dataGridView.Rows[dataGridView.Rows.Count - 1].Cells[i].Value = getSumFromGridColounm(dataGridView, i);
                    }

                    break;
                case "Budget":
                    AddNewRow(dataGridView);
                    dataGridView.Rows[dataGridView.Rows.Count - 1].Cells[0].Value = "Budget Sum";
                    //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
                    for (int i = 1; i < dataGridView.Columns.Count - 1; i++)
                    {
                        double tmp = getBudgetToMonth(selectedYear, (selectedYear == DateTime.Now.Year ? DateTime.Now.Month : 12), dataGridView.Columns[i].Name) - Convert.ToDouble(dataGridView.Rows[13].Cells[i].Value);
                        dataGridView.Rows[dataGridView.Rows.Count - 1].Cells[i].Value = Math.Round(tmp, 2);
                    }
                    break;
                case "Compare":
                    //AddNewRow(dataGridView1);
                    double tmpSum = 0;
                    dataGridView.Rows[dataGridView.Rows.Count - 1].Cells[0].Value = "Budget Sum";
                    //MessageBox.Show("this is datatable row count" + this.dataGridView1.Rows.Count);
                    for (int i = 2; i < 6; i++)
                    {
                        for (int j = 0; j < dataGridView.Rows.Count; j++)
                        {
                            tmpSum += Convert.ToDouble(dataGridView.Rows[j].Cells[i].Value);
                        }
                        dataGridView.Rows[dataGridView.Rows.Count - 1].Cells[i].Value = tmpSum;
                        tmpSum = 0;
                    }
                    break;
            }

        }

        public void ClearDataGrid(DataGridView dataGridView)
        {//清空 dataGridView
            for (int i = dataGridView.Columns.Count - 1; i >= 0; i--)
                dataGridView.Columns.RemoveAt(i);
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
            str = str.Substring(str.Length - 2, 2);
            return str;
        }

        public DataTable LoadFileToDT(String fileName)
        {//将文件存入Datatable
            int i, j;
            fileName = fileName + ".txt";
            DataTable table = new DataTable(fileName);
            fileName = @"Data\" + fileName;
            var tmpList = GetStringArray(fileName);
            var titleList = SplitStr(tmpList[0]);//获取第一列： 列表名

            for (i = 0; i < titleList.Length; i++)
            { //遍历添加列
                table.Columns.Add(titleList[i]);
            }

            //MessageBox.Show("这是数据的列"+tmpList.Count);
            for (i = 1; i < tmpList.Count; i++)//遍历行tmpList.Count
            {
                var tempLine = SplitStr(tmpList[i]);
                var row = table.NewRow();
                for (j = 0; j < titleList.Length && j < tempLine.Length; j++)//遍历列 titleList.Length
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
        public double GetBudget(int Year, string type, string range)//type=quarter/month
        {
            DataTable budgetSheet = LoadFileToDT("budget");
            string sql = "Date='" + Year + "' and 品类='" + type + "'" + " and 分类 = '" + range + "'";
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
        public void AddNewRow(DataGridView dataGridView, String type, string yearMonth)//Type = selected type
        {
            //int index = dataGridView.Rows.Add();
            int index = dataGridView.Rows.Count - 1;
            dataGridView.Rows.Insert(index);

            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                dataGridView.Rows[index].Cells[i].Value = dataGridView.Rows[index + 1].Cells[i].Value;
            }
            if (yearMonth.Contains("-")) {
                dataGridView.Rows[index + 1].Cells["Date"].Value = Convert.ToString(yearMonth);
                dataGridView.Rows[index + 1].Cells["品类"].Value = type;
            }
            else {
                dataGridView.Rows[index + 1].Cells["Date"].Value = Convert.ToString(yearMonth);
                dataGridView.Rows[index + 1].Cells["品类"].Value = type;
            }
        }
        public void AddNewRow(DataGridView dataGridView)
        {
            //int index = dataGridView.Rows.Add();
            int index = dataGridView.Rows.Count - 1;
            dataGridView.Rows.Insert(index);

            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                dataGridView.Rows[index].Cells[i].Value = dataGridView.Rows[index + 1].Cells[i].Value;
            }
        }

        public int LoadYearSelector(ComboBox comboBox)
        {//录入年份
            int currentYear = DateTime.Now.Year;
            if (comboBox.Items.Contains(currentYear.ToString() + "-01"))
            {
                comboBox.Items.Clear();
            }

            if (comboBox.Items.Contains(currentYear.ToString()))
            {
            }
            else
            {
                for (int i = 2018; i <= currentYear; i++)
                {
                    comboBox.Items.Add(Convert.ToString(i));
                }
                comboBox.SelectedItem = currentYear + "";
            }
            return currentYear;
        }
        public String LoadQuarterSelector(ComboBox combox)
        {//录入年份
            int currentYear = DateTime.Now.Year;
            int currentMonth = DateTime.Now.Month;
            if (combox.Items.Contains(currentYear.ToString() + "-01"))
            { }
            else
            {
                for (int i = currentYear; i >= 2018; i--)
                {
                    for (int j = 1; j <= 12; j++)
                    {
                        combox.Items.Add(Convert.ToString(i) + "-" + GetMonth2Dig(j));
                    }
                }
                combox.SelectedItem = currentYear + "-" + currentMonth;
            }
            return combox.SelectedItem.ToString();
        }
        public String GetExpectCell(DataTable table, String sql, string coloumName)
        {//

            DataRow[] tmp = table.Select(sql);
            String result = "";
            if (tmp.Length > 0)
            {
                DataTable tmpTable = tmp.CopyToDataTable();
                result = tmpTable.Rows[0][coloumName].ToString();
                tmpTable.Dispose();
            }

            tmp = null;
            return result;
        }




        public DataTable TransGridToTable(DataGridView dataGridView)
        {
            String name = dataGridView.Name;
            DataTable dataTable = new DataTable(name);

            for (int i = 0; i < dataGridView.Columns.Count; i++)
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

            for (int i = 0; i < tempColumnsCount1; i++)
            {
                dataTable.Columns.Add(dataGridView1.Columns[i].Name);
            }
            for (int i = startColumnGrid2; i < dataGridView2.Columns.Count; i++)
            {
                dataTable.Columns.Add(dataGridView2.Columns[i].Name);
            }

            for (int i = 0; i < dataGridView1.Rows.Count && i < stopLine; i++)//遍历行tmpList.Count
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
            int flag = 0, found = 0;//found为1为找到，0为没找到

            DataTable FileTable = LoadFileToDT(FileName);
            int FilelineCount = FileTable.Rows.Count;
            for (int i = 0; i < GridTable.Rows.Count; i++)
            {//遍历想插入的表 GridTable的第i行
                for (int j = 0; j < FilelineCount && found == 0; j++)
                {//遍历源文件的行 FileTable的第j行
                    for (int k = 0; k < index; k++)
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
                if (found == 0)
                { //如果没找到相同的行
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
            StreamWriter wr = new StreamWriter(F, Encoding.GetEncoding(ACCode));
            String tmpStr = "";

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                tmpStr = tmpStr + dataTable.Columns[i].ColumnName.ToString() + ",";
            }
            tmpStr = tmpStr.Substring(0, tmpStr.Length - 1);
            wr.WriteLine(tmpStr, Encoding.GetEncoding(ACCode));
            tmpStr = "";

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    if (dataTable.Rows[i][j].ToString() != "")
                    {
                        tmpStr = tmpStr + dataTable.Rows[i][j].ToString() + ",";
                    }
                }
                if (tmpStr != "")
                {
                    tmpStr = tmpStr.Substring(0, tmpStr.Length - 1);
                    wr.WriteLine(tmpStr, Encoding.GetEncoding(ACCode));
                    tmpStr = "";
                }
            }
            wr.Close();
            F.Close();

        }
        public double getSumByTime(int year, int month, string type)
        {
            DataTable fileMonth = LoadFileToDT("month");
            double sum = 0;
            string sql = "Date<='" + year + "." + GetMonth2Dig(month) + "' and Date>='" + year + ".00'";
            if (fileMonth.Select(sql).Length > 0)
            {
                DataTable temp = new DataTable();
                temp = fileMonth.Select(sql).CopyToDataTable();
                for (int i = 0; i < temp.Rows.Count; i++)
                {
                    sum = sum + Convert.ToDouble(temp.Rows[i][type]);

                }
                temp = null;
            }
            return sum;
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
                        if (FileTable.Rows[j][k].ToString() != GridTable.Rows[i][k].ToString() )
                        {
                            //MessageBox.Show("not equal");
                            flag = 1;//如果两个表目标行的前index格不相等，清空
                            break;
                        }else if (FileTable.Rows[j][k].ToString()==null)
                        {
                            flag = 0;
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

        public string[] getTypeList(string fileName,int start, int stop)
        {
            DataTable file = LoadFileToDT(fileName);
            string list="";
            for (int i = start;i< file.Columns.Count && i<stop; i++)
            {
                    list += file.Columns[i].ColumnName + " ";

            }
            list = list.TrimEnd();
            return list.Split(' ');
        }
        public string[] getTypeList(string fileName,string lineName)
        {
            DataTable distinct = new DataView(LoadFileToDT(fileName)).ToTable(true, lineName);
            
            string[] list = new string[distinct.Rows.Count];
            for(int i = 0; i < distinct.Rows.Count; i++)
            {
                list[i] = distinct.Rows[i][0].ToString();
            }
            return list;
        }
    }
}