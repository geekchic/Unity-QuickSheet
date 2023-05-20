///////////////////////////////////////////////////////////////////////////////
///
/// ExcelMachineEditor.cs
///
/// (c)2014 Kim, Hyoun Woo
///
///////////////////////////////////////////////////////////////////////////////
using UnityEngine;
using UnityEditor;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;

namespace UnityQuickSheet
{
    /// <summary>
    /// Custom editor script class for excel file setting.
    /// </summary>
    [CustomEditor(typeof(ExcelMachine))]
    public class ExcelMachineEditor : BaseMachineEditor
    {
        protected override void OnEnable()
        {
            base.OnEnable();

            machine = target as ExcelMachine;
            if (machine != null && ExcelSettings.Instance != null)
            {
                if (string.IsNullOrEmpty(ExcelSettings.Instance.RuntimePath) == false)
                    machine.RuntimeClassPath = ExcelSettings.Instance.RuntimePath;
                if (string.IsNullOrEmpty(ExcelSettings.Instance.EditorPath) == false)
                    machine.EditorClassPath = ExcelSettings.Instance.EditorPath;
                if (string.IsNullOrEmpty(ExcelSettings.Instance.ScriptableObjectPath) == false)
                    machine.ScriptableObjectPath = ExcelSettings.Instance.ScriptableObjectPath;
            }
        }

        public override void OnInspectorGUI()
        {
            base.OnInspectorGUI();

            ExcelMachine machine = target as ExcelMachine;

            GUILayout.Label("Excel Spreadsheet Settings:", headerStyle);

            GUILayout.BeginHorizontal();
            GUILayout.Label("File:", GUILayout.Width(50));

            string path = string.Empty;
            if (string.IsNullOrEmpty(machine.excelFilePath))
                path = Application.dataPath;
            else
                path = machine.excelFilePath;

            machine.excelFilePath = GUILayout.TextField(path, GUILayout.Width(250));
            if (GUILayout.Button("...", GUILayout.Width(20)))
            {
                string folder = Path.GetDirectoryName(path);
#if UNITY_EDITOR_WIN
                path = EditorUtility.OpenFilePanel("Open Excel file", folder, "excel files;*.xls;*.xlsx");
#else // for UNITY_EDITOR_OSX
                path = EditorUtility.OpenFilePanel("Open Excel file", folder, "xls");
#endif
                if (path.Length != 0)
                {
                    machine.SpreadSheetName = Path.GetFileName(path);

                    // the path should be relative not absolute one to make it work on any platform.
                    int index = path.IndexOf("Assets");
                    if (index >= 0)
                    {
                        // set relative path
                        machine.excelFilePath = path.Substring(index);

                        // pass absolute path
                        machine.SheetNames = new ExcelQuery(path).GetSheetNames();
                    }
                    else
                    {
                        EditorUtility.DisplayDialog("Error",
                            @"Wrong folder is selected.
                        Set a folder under the 'Assets' folder! \n
                        The excel file should be anywhere under  the 'Assets' folder", "OK");
                        return;
                    }
                }
            }
            GUILayout.EndHorizontal();

            // Failed to get sheet name so we just return not to make editor on going.
            if (machine.SheetNames.Length == 0)
            {
                EditorGUILayout.Separator();
                EditorGUILayout.LabelField("Error: Failed to retrieve the specified excel file.");
                EditorGUILayout.LabelField("If the excel file is opened, close it then reopen it again.");
                return;
            }

            // spreadsheet name should be read-only
            EditorGUILayout.TextField("Spreadsheet File: ", machine.SpreadSheetName);

            EditorGUILayout.Separator();

            if (GUILayout.Button("Generate Enums"))
            {
                GenerateEnums();
            }

            EditorGUILayout.Space();

            using (new GUILayout.HorizontalScope())
            {
                EditorGUILayout.LabelField("Worksheet: ", GUILayout.Width(100));
                bool dirty = false;
                var sheetIndex = EditorGUILayout.Popup(machine.CurrentSheetIndex, machine.SheetNames);
                if (sheetIndex != machine.CurrentSheetIndex)
                {
                    dirty = true;
                }

                machine.CurrentSheetIndex = sheetIndex;

                if (machine.SheetNames != null)
                    machine.WorkSheetName = machine.SheetNames[machine.CurrentSheetIndex];

                if (dirty)
                {
                    Import(true);
                }

                if (GUILayout.Button("Refresh", GUILayout.Width(60)))
                {
                    // reopen the excel file e.g) new worksheet is added so need to reopen.
                    machine.SheetNames = new ExcelQuery(machine.excelFilePath).GetSheetNames();

                    // one of worksheet was removed, so reset the selected worksheet index
                    // to prevent the index out of range error.
                    if (machine.SheetNames.Length <= machine.CurrentSheetIndex)
                    {
                        machine.CurrentSheetIndex = 0;

                        string message = "Worksheet was changed. Check the 'Worksheet' and 'Update' it again if it is necessary.";
                        EditorUtility.DisplayDialog("Info", message, "OK");
                    }
                }
            }

            EditorGUILayout.Separator();

            GUILayout.BeginHorizontal();

            if (machine.HasColumnHeader())
            {
                if (GUILayout.Button("Update"))
                    Import();
                if (GUILayout.Button("Reimport"))
                    Import(true);
            }
            else
            {
                if (GUILayout.Button("Import"))
                    Import();
            }

            GUILayout.EndHorizontal();

            EditorGUILayout.Separator();

            DrawHeaderSetting(machine);

            EditorGUILayout.Separator();

            GUILayout.Label("Path Settings:", headerStyle);

            machine.TemplatePath = EditorGUILayout.TextField("Template: ", machine.TemplatePath);
            machine.RuntimeClassPath = EditorGUILayout.TextField("Runtime: ", machine.RuntimeClassPath);
            machine.EditorClassPath = EditorGUILayout.TextField("Editor:", machine.EditorClassPath);
            machine.ScriptableObjectPath = EditorGUILayout.TextField("Data:", machine.ScriptableObjectPath);

            machine.onlyCreateDataClass = EditorGUILayout.Toggle("Only DataClass", machine.onlyCreateDataClass);

            EditorGUILayout.Separator();

            if (GUILayout.Button("Generate"))
            {
                if (string.IsNullOrEmpty(machine.SpreadSheetName) || string.IsNullOrEmpty(machine.WorkSheetName))
                {
                    Debug.LogWarning("No spreadsheet or worksheet is specified.");
                    return;
                }

	            Directory.CreateDirectory(Application.dataPath + Path.DirectorySeparatorChar + machine.RuntimeClassPath);
	            Directory.CreateDirectory(Application.dataPath + Path.DirectorySeparatorChar + machine.EditorClassPath);

                ScriptPrescription sp = Generate(machine);
                if (sp != null)
                {
                    Debug.Log("Successfully generated!");
                }
                else
                    Debug.LogError("Failed to create a script from excel.");
            }

            if (GUI.changed)
            {
                EditorUtility.SetDirty(machine);
            }
        }

        /// <summary>
        /// Import the specified excel file and prepare to set type of each cell.
        /// </summary>
        protected override void Import(bool reimport = false)
        {
            ExcelMachine machine = target as ExcelMachine;

            string path = machine.excelFilePath;
            string sheet = machine.WorkSheetName;

            if (string.IsNullOrEmpty(path))
            {
                string msg = "You should specify spreadsheet file first!";
                EditorUtility.DisplayDialog("Error", msg, "OK");
                return;
            }

            if (!File.Exists(path))
            {
                string msg = string.Format("File at {0} does not exist.",path);
                EditorUtility.DisplayDialog("Error", msg, "OK");
                return;
            }

            string error = string.Empty;
            var titles = new ExcelQuery(path, sheet).GetTitle(ref error);
            if (titles == null || !string.IsNullOrEmpty(error))
            {
                EditorUtility.DisplayDialog("Error", error, "OK");
                return;
            }
            else
            {
                // check the column header is valid
                foreach(string column in titles)
                {
                    if (!IsValidHeader(column))
                    {
                        error = string.Format(@"Invalid column header name {0}. Any c# keyword should not be used for column header. Note it is not case sensitive.", column);
                        EditorUtility.DisplayDialog("Error", error, "OK");
                        return;
                    }
                }
            }

            List<string> titleList = titles.ToList();
            {
                if (machine.HasColumnHeader() && reimport == false)
                {
                    var headerDic = machine.ColumnHeaderList.ToDictionary(header => header.name);

                    // collect non-changed column headers
                    var exist = titleList.Select(t => GetColumnHeaderString(t))
                        .Where(e => headerDic.ContainsKey(e) == true)
                        .Select(t => new ColumnHeader { name = t, type = headerDic[t].type, isArray = headerDic[t].isArray, OrderNO = headerDic[t].OrderNO });


                    // collect newly added or changed column headers
                    var changed = titleList.Select(t => GetColumnHeaderString(t))
                        .Where(e => headerDic.ContainsKey(e) == false)
                        .Select(t => ParseColumnHeader(t, titleList.IndexOf(t)));

                    // merge two list via LINQ
                    var merged = exist.Union(changed).OrderBy(x => x.OrderNO);

                    machine.ColumnHeaderList.Clear();
                    machine.ColumnHeaderList = merged.ToList();
                }
                else
                {
                    machine.ColumnHeaderList.Clear();
                    if (titleList.Count > 0)
                    {
                        int order = 0;
                        machine.ColumnHeaderList = titleList.Select(e => ParseColumnHeader(e, order++)).ToList();
                    }
                    else
                    {
                        string msg = string.Format("An empty workhheet: [{0}] ", sheet);
                        Debug.LogWarning(msg);
                    }
                }
            }

            EditorUtility.SetDirty(machine);
            AssetDatabase.SaveAssets();
        }

        /// <summary>
        /// Generate AssetPostprocessor editor script file.
        /// </summary>
        protected override void CreateAssetCreationScript(BaseMachine m, ScriptPrescription sp)
        {
            ExcelMachine machine = target as ExcelMachine;

            sp.className = machine.WorkSheetName;
            sp.dataClassName = machine.WorkSheetName + "Data";
            sp.worksheetClassName = machine.WorkSheetName;

            // where the imported excel file is.
            sp.importedFilePath = machine.excelFilePath;

            // path where the .asset file will be created.
            //string path = Path.GetDirectoryName(machine.excelFilePath);
            string path = Path.Combine("Assets", machine.ScriptableObjectPath, machine.WorkSheetName + ".asset");
            sp.assetFilepath = path.Replace('\\', '/');
            sp.assetPostprocessorClass = machine.WorkSheetName + "AssetPostprocessor";
            sp.template = GetTemplate("PostProcessor");

            // write a script to the given folder.
            using (var writer = new StreamWriter(TargetPathForAssetPostProcessorFile(machine.WorkSheetName)))
            {
                writer.Write(new ScriptGenerator(sp).ToString());
                writer.Close();
            }
        }

        private void GenerateEnums()
        {
            ExcelMachine machine = target as ExcelMachine;

            string path = machine.excelFilePath;

            if (string.IsNullOrEmpty(path))
            {
                string msg = "You should specify spreadsheet file first!";
                EditorUtility.DisplayDialog("Error", msg, "OK");
                return;
            }

            if (!File.Exists(path))
            {
                string msg = string.Format("File at {0} does not exist.", path);
                EditorUtility.DisplayDialog("Error", msg, "OK");
                return;
            }

            var excelQuery = new ExcelQuery(path);
            var enumTables = excelQuery.GetEnumTables();
            if (enumTables == null)
            {
                EditorUtility.DisplayDialog("Error", "Not Found Enums", "OK");
                return;
            }
            else if (enumTables.Length == 0)
            {
                EditorUtility.DisplayDialog("Error", "Not Found Enums", "OK");
                return;
            }
            else
            {
                Directory.CreateDirectory(Path.Combine(Application.dataPath, machine.RuntimeClassPath));

                ScriptPrescription sp = GenerateEnums(machine, excelQuery);
                if (sp != null)
                {
                    Debug.Log("Successfully generated!");
                }
                else
                    Debug.LogError("Failed to create a script from excel.");
            }
        }


        /// <summary>
        /// Generate script files with the given templates.
        /// Total four files are generated, two for runtime and others for editor.
        /// </summary>
        private ScriptPrescription GenerateEnums(BaseMachine m, ExcelQuery excelQuery)
        {
            if (m == null)
                return null;

            ScriptPrescription sp = new ScriptPrescription();
            CreateEnumsClassScript(m, excelQuery, sp);
            AssetDatabase.Refresh();
            return sp;
        }

        /// <summary>
        /// Create a data class which describes the spreadsheet and write it down on the specified folder.
        /// </summary>
        private void CreateEnumsClassScript(BaseMachine machine, ExcelQuery excelQuery, ScriptPrescription sp)
        {
            // check the directory path exists
            string fullPath = Path.Combine(Application.dataPath, machine.RuntimeClassPath,"Enums.cs");
            string folderPath = Path.GetDirectoryName(fullPath);
            if (!Directory.Exists(folderPath))
            {
                EditorUtility.DisplayDialog(
                    "Warning",
                    "The folder for runtime script files does not exist. Check the path " + folderPath + " exists.",
                    "OK"
                    );
                return;
            }

            List<EnumTableData> enumTables = new List<EnumTableData>();


            var tables = excelQuery.GetEnumTables();

            foreach (var table in tables)
            {
                var data = new EnumTableData();
                data.Name = table.Name;

                var sheet = table.GetXSSFSheet();
                int startRow = table.GetStartCellReference().Row + 1;
                int endRow = table.GetEndCellReference().Row;
                int startCol = table.GetStartCellReference().Col;
                int endCol = table.GetEndCellReference().Col;

                int index = 0;
                for (int r = startRow; r <= endRow; r++)
                {
                    IRow row = sheet.GetRow(r);
                    string key = row.GetCell(startCol).StringCellValue;
                    int value = (row.Count() > 1) ? Convert.ToInt32(row.GetCell(startCol + 1).NumericCellValue) : index;
                    data.MemberFields.Add(key, value);
                    index++;
                }

                enumTables.Add(data);
            }

            sp.className = "Enums";
            sp.template = GetTemplate("EnumsClass");
            sp.enumTables = enumTables.ToArray();

            // write a script to the given folder.		
            using (var writer = new StreamWriter(fullPath))
            {
                writer.Write(new ScriptGenerator(sp).ToString());
                writer.Close();
            }
        }
    }
}