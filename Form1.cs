using OctopusCore;
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

namespace OctopusTools
{
    public partial class Form1 : Form
    {
        private VariableService _variableService;
        public Form1()
        {
            InitializeComponent();
            _variableService = new VariableService();
            GetProjectNames();
            string name = comboBox1.Text;
            GetScopes(name);
        }

        private void GetProjectNames()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            var projectNames = _variableService.GetProjects();
            foreach (var name in projectNames)
            {
                DataRow row = dt.Rows.Add();
                row[0] = name;
            }

            comboBox1.ValueMember = "Name";
            comboBox1.DisplayMember = "Name";
            comboBox1.DataSource = dt;
        }

        private void GetScopes(string name)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            var scopeNames = _variableService.GetScopes(name);
            DataRow firstRow = dt.Rows.Add();
            firstRow[0] = "All";
            foreach (var scopeName in scopeNames)
            {
                DataRow row = dt.Rows.Add();
                row[0] = scopeName;
            }

            comboBox2.ValueMember = "Name";
            comboBox2.DisplayMember = "Name";
            comboBox2.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel|*.xlsx";
            saveFileDialog1.Title = "Save file";
            saveFileDialog1.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.FileName != "" && Path.GetExtension(saveFileDialog1.FileName) == ".xlsx")
            {
                _variableService.Export(comboBox1.Text, comboBox2.Text, saveFileDialog1.FileName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var selectFileDialog = new OpenFileDialog())
            {
                if (selectFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _variableService.Import(comboBox1.Text, selectFileDialog.FileName);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var selectFileDialog = new OpenFileDialog())
            {
                if (selectFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var tbl = _variableService.Compare(comboBox1.Text, selectFileDialog.FileName);

                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Excel|*.xlsx";
                    saveFileDialog1.Title = "Save file";
                    saveFileDialog1.ShowDialog();

                    // If the file name is not an empty string open it for saving.
                    if (saveFileDialog1.FileName != "" && Path.GetExtension(saveFileDialog1.FileName) == ".xlsx")
                    {
                        _variableService.ExportCompare(saveFileDialog1.FileName, comboBox1.Text, tbl);
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (var selectFileDialog = new OpenFileDialog())
            {
                selectFileDialog.Multiselect = true;
                if (selectFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (selectFileDialog.FileNames.Count() != 2)
                    {
                        return;
                    }

                    var tbl1 = _variableService.Compare(comboBox1.Text, selectFileDialog.FileNames[0]);
                    var tbl2 = _variableService.Compare(comboBox1.Text, selectFileDialog.FileNames[1]);

                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Excel|*.xlsx";
                    saveFileDialog1.Title = "Save file";
                    saveFileDialog1.ShowDialog();

                    // If the file name is not an empty string open it for saving.
                    if (saveFileDialog1.FileName != "" && Path.GetExtension(saveFileDialog1.FileName) == ".xlsx")
                    {
                        _variableService.ExportCompareTwoExcelFile(saveFileDialog1.FileName, comboBox1.Text, tbl1, tbl2);
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Excel|*.xlsx";
            saveFileDialog1.Title = "Save file";
            saveFileDialog1.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (saveFileDialog1.FileName != "" && Path.GetExtension(saveFileDialog1.FileName) == ".xlsx")
            {
                _variableService.ExportSeparatedEnvironments(comboBox1.Text, comboBox2.Text, saveFileDialog1.FileName);
            }
        }
    }
}
