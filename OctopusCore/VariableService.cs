using Octopus.Client;
using Octopus.Client.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OctopusCore
{
    public class VariableService
    {
        private OctopusRepository _octopusRepository;
        public VariableService()
        {
            var endPoint = ConfigurationManager.AppSettings["Octopus.EndPoint"];
            var apiKey = ConfigurationManager.AppSettings["Octopus.ApiKey"];
            _octopusRepository = new OctopusRepository(new OctopusServerEndpoint(endPoint, apiKey));
        }

        public List<string> GetProjects()
        {
            return _octopusRepository.Projects.GetAll().Select(p => p.Name).ToList();
        }

        public List<string> GetScopes(string name)
        {
            var project = _octopusRepository.Projects.FindByName(name);
            var variableSet = _octopusRepository.VariableSets.Get(project.Link("Variables"));
            return variableSet.ScopeValues.Environments.Select(s => s.Name).ToList();
        }

        public void Export(string projectName, string scope, string fileName)
        {
            // Find the project that owns the variables we want to edit
            var project = _octopusRepository.Projects.FindByName(projectName);

            var variableSet = _octopusRepository.VariableSets.Get(project.Link("Variables"));
            var scopeValue = variableSet.ScopeValues.Environments;

            //save to excel
            FileInfo newFile = new FileInfo(fileName);
            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet ws = xlPackage.Workbook.Worksheets.Add("octopus");
                ws.Cells[1, 1].Value = "Name";
                ws.Cells[1, 2].Value = "Value";
                ws.Cells[1, 3].Value = "Scope";
                int start = 2;
                foreach (var variable in variableSet.Variables)
                {
                    var variableScope = variable.Scope.Values.FirstOrDefault();

                    var scopes = scopeValue.Where(c => variableScope != null && variableScope.Any(v => v == c.Id)).ToList();
                    var names = scopes.Select(c => c.Name);
                    if (scope == "All" || names.Contains(scope) || variableScope == null)
                    {
                        ws.Cells[start, 1].Value = variable.Name;
                        ws.Cells[start, 2].Value = variable.Value;
                        ws.Cells[start, 3].Value = string.Join(",", names);
                        start++;
                    }
                }

                xlPackage.Save();
            }
        }

        public void Import(string projectName, string fileName)
        {
            // Find the project that owns the variables we want to edit
            var project = _octopusRepository.Projects.FindByName(projectName);

            var variableSet = _octopusRepository.VariableSets.Get(project.Link("Variables"));
            var scopeValue = variableSet.ScopeValues.Environments;

            //save to excel
            DataTable tbl = new DataTable();

            using (var pck = new ExcelPackage())
            {
                using (var stream = File.OpenRead(fileName))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.FirstOrDefault();
                var headerRow = 1;

                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    tbl.Columns.Add(ws.Cells[headerRow, col].Text ?? string.Format("Column {0}", col));
                }
                for (int rowNum = headerRow + 1; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    DataRow row = tbl.Rows.Add();
                    for (int col = 1; col <= ws.Dimension.End.Column; col++)
                    {
                        row[col - 1] = ws.Cells[rowNum, col].Text ?? "";
                    }
                }
            }

            foreach (DataRow row in tbl.Rows)
            {
                if (string.IsNullOrEmpty(row["Scope"].ToString()))
                {
                    continue;
                }

                var scopes = row["Scope"].ToString().Split(',');
                var scopeIds = scopeValue.FindAll(s => scopes.Contains(s.Name)).Select(s => s.Id).ToArray();

                var variableName = row["Name"].ToString();
                var variable = variableSet.Variables.FirstOrDefault(v => v.Name == variableName && v.Scope.Values.Any() && v.Scope.Values.FirstOrDefault().Intersect(scopeIds).Count() == scopeIds.Count());

                if (variable == null)
                {
                    variableSet.Variables.Add(new VariableResource()
                    {
                        Name = row["Name"].ToString(),
                        Value = row["Value"].ToString(),
                        Scope = new ScopeSpecification()
                        {
                            // Scope the variable to two environments using their environment ID
                            { ScopeField.Environment, new ScopeValue(scopeIds)}
                        },
                    });
                }
                else if (!variable.IsSensitive)
                {
                    var value = row["Value"].ToString();
                    variable.Value = value;
                }
            }

            // Save the variables
            _octopusRepository.VariableSets.Modify(variableSet);
        }

        public DataTable Compare(string projectName, string fileName)
        {
            //save to excel
            DataTable tbl = new DataTable();

            using (var pck = new ExcelPackage())
            {
                using (var stream = File.OpenRead(fileName))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.FirstOrDefault();
                var headerRow = 1;

                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    tbl.Columns.Add(ws.Cells[headerRow, col].Text ?? string.Format("Column {0}", col));
                }
                for (int rowNum = headerRow + 1; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    DataRow row = tbl.Rows.Add();
                    for (int col = 1; col <= ws.Dimension.End.Column; col++)
                    {
                        row[col - 1] = ws.Cells[rowNum, col].Text ?? "";
                    }
                }
            }

            return tbl;
        }

        public void ExportCompare(string fileName, string projectName, DataTable tbl)
        {
            // Find the project that owns the variables we want to edit
            var project = _octopusRepository.Projects.FindByName(projectName);

            var variableSet = _octopusRepository.VariableSets.Get(project.Link("Variables"));
            var scopeValue = variableSet.ScopeValues.Environments;

            FileInfo newFile = new FileInfo(fileName);
            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet ws = xlPackage.Workbook.Worksheets.Add("octopus");

                //Red
                var conditionalFormattingRule05 = ws.ConditionalFormatting.AddExpression(ws.Cells[2, 8, tbl.Rows.Count + 1, 8]);
                conditionalFormattingRule05.Formula = "($H2=FALSE)";
                conditionalFormattingRule05.Style.Fill.PatternType = ExcelFillStyle.Solid;
                conditionalFormattingRule05.Style.Fill.BackgroundColor.Color = Color.FromArgb(234, 153, 153);

                ws.Cells[1, 1].Value = "Name";
                ws.Cells[1, 2].Value = "Value";
                ws.Cells[1, 3].Value = "NewValue";
                ws.Cells[1, 4].Value = "FileScope";
                ws.Cells[1, 5].Value = "VariableScope";
                ws.Cells[1, 6].Value = "Status";

                ws.Cells[1, 8].Value = "CompareValue";

                int start = 2;

                foreach (DataRow row in tbl.Rows)
                {
                    var variableName = row["Name"].ToString();
                    var scopes = row["Scope"].ToString().Split(',');

                    ws.Cells[start, 1].Value = variableName;
                    ws.Cells[start, 4].Value = scopes;
                    ws.Cells[start, 8].Formula = $"=IF(B{start}=C{start},TRUE,FALSE)";

                    if (string.IsNullOrEmpty(row["Scope"].ToString()))
                    {
                        ws.Cells[start, 6].Value = "SKIP";
                        start++;

                        continue;
                    }

                    var scopeIds = scopeValue.FindAll(s => scopes.Contains(s.Name)).Select(s => s.Id).ToArray();

                    var variable = variableSet.Variables.FirstOrDefault(v => v.Name == variableName && v.Scope.Values.Any() && v.Scope.Values.FirstOrDefault().Intersect(scopeIds).Count() == scopeIds.Count());

                    if (variable == null)
                    {
                        ws.Cells[start, 3].Value = row["Value"].ToString();
                        ws.Cells[start, 6].Value = "ADD";

                        //variableSet.Variables.Add(new VariableResource()
                        //{
                        //    Name = row["Name"].ToString(),
                        //    Value = row["Value"].ToString(),
                        //    Scope = new ScopeSpecification()
                        //{
                        //    // Scope the variable to two environments using their environment ID
                        //    { ScopeField.Environment, new ScopeValue(scopeIds)}
                        //},
                        //});
                    }
                    else if (!variable.IsSensitive)
                    {
                        ws.Cells[start, 2].Value = variable.Value;
                        ws.Cells[start, 3].Value = row["Value"].ToString();
                        ws.Cells[start, 5].Value = variable.Scope.Values;
                        ws.Cells[start, 6].Value = "UPDATE";

                        var value = row["Value"].ToString();
                        variable.Value = value;
                    }
                    else
                    {
                        ws.Cells[start, 6].Value = "SKIP";
                    }

                    start++;
                }

                xlPackage.Save();
            }
        }

        public void ExportCompareTwoExcelFile(string fileName, string projectName, DataTable tbl1, DataTable tbl2)
        {
            FileInfo newFile = new FileInfo(fileName);
            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet ws = xlPackage.Workbook.Worksheets.Add("octopus");

                var tbl1NameScopes = tbl1.AsEnumerable().Select(x => new { Name = x["Name"].ToString(), Value = x["Value"].ToString(), Scope = x["Scope"].ToString() });
                var tbl2NameScopes = tbl2.AsEnumerable().Select(x => new { Name = x["Name"].ToString(), Value = x["Value"].ToString(), Scope = x["Scope"].ToString() });

                var updates = tbl1NameScopes.Where(x => tbl2NameScopes.Any(y => y.Name == x.Name && y.Scope == x.Scope)).ToList();
                var deletes = tbl1NameScopes.Where(x => !tbl2NameScopes.Any(y => y.Name == x.Name && y.Scope == x.Scope)).ToList();
                var adds = tbl2NameScopes.Where(x => !tbl1NameScopes.Any(y => y.Name == x.Name && y.Scope == x.Scope)).ToList();

                // Red
                var max = updates.Count + deletes.Count + adds.Count;
                var conditionalFormattingRule05 = ws.ConditionalFormatting.AddExpression(ws.Cells[2, 8, max + 1, 8]);
                conditionalFormattingRule05.Formula = "($H2=FALSE)";
                conditionalFormattingRule05.Style.Fill.PatternType = ExcelFillStyle.Solid;
                conditionalFormattingRule05.Style.Fill.BackgroundColor.Color = Color.FromArgb(234, 153, 153);

                ws.Cells[1, 1].Value = "Name";
                ws.Cells[1, 2].Value = "Value";
                ws.Cells[1, 3].Value = "NewValue";
                ws.Cells[1, 4].Value = "FileOneScope";
                ws.Cells[1, 5].Value = "FileTwoScope";
                ws.Cells[1, 6].Value = "Status";

                ws.Cells[1, 8].Value = "CompareValue";

                int start = 2;

                foreach (var update in updates)
                {
                    ws.Cells[start, 1].Value = update.Name;
                    ws.Cells[start, 2].Value = update.Value;
                    ws.Cells[start, 4].Value = update.Scope;
                    ws.Cells[start, 8].Formula = $"=IF(B{start}=C{start},TRUE,FALSE)";

                    var tbl2Data = tbl2NameScopes.Where(x => x.Name == update.Name && x.Scope == update.Scope).FirstOrDefault();

                    ws.Cells[start, 3].Value = tbl2Data?.Value;
                    ws.Cells[start, 5].Value = tbl2Data?.Scope;
                    ws.Cells[start, 6].Value = "UPDATE";

                    start++;
                }

                foreach (var delete in deletes)
                {
                    ws.Cells[start, 1].Value = delete.Name;
                    ws.Cells[start, 2].Value = delete.Value;
                    ws.Cells[start, 4].Value = delete.Scope;
                    ws.Cells[start, 8].Formula = "=FALSE";

                    ws.Cells[start, 6].Value = "DELETE";

                    start++;
                }

                foreach (var add in adds)
                {
                    ws.Cells[start, 1].Value = add.Name;
                    ws.Cells[start, 3].Value = add.Value;
                    ws.Cells[start, 5].Value = add.Scope;
                    ws.Cells[start, 8].Formula = "=FALSE";

                    ws.Cells[start, 6].Value = "ADD";

                    start++;
                }

                xlPackage.Save();
            }
        }

        public void ExportSeparatedEnvironments(string projectName, string scope, string fileName)
        {
            // Find the project that owns the variables we want to edit
            var project = _octopusRepository.Projects.FindByName(projectName);

            var variableSet = _octopusRepository.VariableSets.Get(project.Link("Variables"));
            var scopeValue = variableSet.ScopeValues.Environments;

            //save to excel
            FileInfo newFile = new FileInfo(fileName);
            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
            {
                ExcelWorksheet ws = xlPackage.Workbook.Worksheets.Add("octopus");
                ws.Cells[1, 1].Value = "Name";
                ws.Cells[1, 2].Value = "Value";
                ws.Cells[1, 3].Value = "Scope";
                int start = 2;
                foreach (var variable in variableSet.Variables)
                {
                    var variableScope = variable.Scope.Values.FirstOrDefault();

                    var scopes = scopeValue.Where(c => variableScope != null && variableScope.Any(v => v == c.Id)).ToList();
                    var names = scopes.Select(c => c.Name);
                    if (scope == "All" || names.Contains(scope) || variableScope == null)
                    {
                        foreach (var name in names)
                        {
                            ws.Cells[start, 1].Value = variable.Name;
                            ws.Cells[start, 2].Value = variable.Value;
                            ws.Cells[start, 3].Value = name;
                            start++;
                        }
                    }
                }

                xlPackage.Save();
            }
        }
    }
}

