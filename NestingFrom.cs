using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using 激光报表;

namespace 报表工具
{
    public partial class NestingFrom : Form
    {
        string contractNumber;
        object[,] dataArray;
        int rows;
        Worksheet product;
        Range selectionRange;

        object[] ts;//厚度列表
        System.Data.DataTable dataTable;

        public NestingFrom(List<Product> products)
        {
            InitializeComponent();
            this.dataTable = ToolsHelper.CreateCsvTable(products);
            this.ts = ToolsHelper.GetUniqueTValues(dataTable);
        }
        //public NestingFrom(string contractNumber, object[,] dataArray, Worksheet product, Range selectionRange)
        //{
        //    InitializeComponent();

        //    this.contractNumber = contractNumber;
        //    this.dataArray = dataArray;
        //    this.product = product;
        //    this.selectionRange = selectionRange;

        //    List<object[]> filteredData;
        //    rows = selectionRange.Rows.Count;

        //    //如果选中了多行，说明用户想要套料多个订单，获取合同编号和物料号进行筛选
        //    if (rows > 1)
        //    {
        //        this.checkBox1.Checked = true;
        //        List<ContractMaterialData> contractMaterialDatas = ToolsHelper.GetContractAndMaterialFromSelection(selectionRange);
        //        filteredData = ExcelHelper.GetFilteredRowsByContractAndMaterial(contractMaterialDatas, dataArray);
        //    }
        //    else {
        //        filteredData = ExcelHelper.GetFilteredRowsByContractNumber(contractNumber, dataArray);
        //    }

        //    System.Data.DataTable dataTable = ToolsHelper.CteateCsvTable(filteredData, product);
        //    //dataTable = ToolsHelper.CteateCsvTable(filteredData, product);
        //    //获取厚度列表
        //    this.dataTable = dataTable;
        //    this.ts = ToolsHelper.GetUniqueTValues(dataTable);
        //}
        private void Okbutton_Click(object sender, EventArgs e)
        {
            //保存表单数据
            Properties.Settings.Default.width = int.Parse(this.widthTextBox.Text);
            Properties.Settings.Default.maxLength = int.Parse(this.maxLengthTextBox.Text);
            Properties.Settings.Default.minLength = int.Parse(this.minLengthTextBox.Text);
            Properties.Settings.Default.maxWidth = int.Parse(this.maxWidthTextBox.Text);
            Properties.Settings.Default.minWidth = int.Parse(this.minWidthTextBox.Text);

            if (!checkBox1.Checked)
            {

            }
            if (this.comboBox1.SelectedItem == null) return;
            //筛选厚度
            dataTable = ToolsHelper.ApplyQuantityRule(dataTable, comboBox1.Text);
            //生成板材数据表
            System.Data.DataTable sheetDataTable =
                ToolsHelper.CreateRectangleTable(Properties.Settings.Default.maxLength, Properties.Settings.Default.minLength,
                Properties.Settings.Default.maxWidth, Properties.Settings.Default.minWidth, 4);
            //合并订单数据表和板材数据表
            dataTable.Merge(sheetDataTable);
            ExcelHelper.StartNesting(dataTable, contractNumber, "D:\\排版\\CSV");
            this.Close();
        }

        private void NestingFrom_Load(object sender, EventArgs e)
        {
            this.widthTextBox.Text = Properties.Settings.Default.width.ToString();
            this.maxLengthTextBox.Text = Properties.Settings.Default.maxLength.ToString();
            this.minLengthTextBox.Text = Properties.Settings.Default.minLength.ToString();
            this.maxWidthTextBox.Text = Properties.Settings.Default.maxWidth.ToString();
            this.minWidthTextBox.Text = Properties.Settings.Default.minWidth.ToString();
            this.contractNumberTextBox.Text = contractNumber;
            this.comboBox1.Items.AddRange(ts);
            this.comboBox1.SelectedIndex = 0;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                this.contractNumberTextBox.Enabled = false;
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}
