using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Reflection;
using UnityEditor;
using UnityEngine;
using UnityQuickSheet;

///
/// !!! Machine generated code !!!
///
[CustomEditor(typeof(Item))]
public class ItemEditor : BaseGoogleEditor<Item>
{	    
    public override bool Load()
    {        
        Item targetData = target as Item;
        
        List<ItemData> myDataList = new List<ItemData>();

        var t = typeof(ItemData);
        PropertyInfo[] p = t.GetProperties();

        var service = GoogleDataSettings.Instance.Service;
        var spreadsheetId = targetData.SheetName;
        var sheetNameAndRange = $"{targetData.WorksheetName}";
        SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, sheetNameAndRange);
        ValueRange response = request.Execute();
        var rows = response.Values;

        for (int i = 1; i < rows.Count; i++)
        {
            var cells = rows[i];
            var item = (ItemData)Activator.CreateInstance(t);
            for (var j = 0; j < p.Length; j++)
            {
                var cell = cells[j];

                if (cell == null)  // skip empty cell
                    continue;

                var property = p[j];
                if (property.CanWrite)
                {
                    try
                    {
                        var value = ConvertFrom(cell, property.PropertyType);
                        property.SetValue(item, value, null);
                    }
                    catch (Exception e)
                    {
                        string pos = string.Format("Row[{0}], Cell[{1}]", i, j);
                        Debug.LogError(string.Format("Excel File {0} Deserialize Exception: {1} at {2}", targetData.WorksheetName, e.Message, pos));
                    }
                }
            }

            myDataList.Add(item);
        }

        targetData.dataArray = myDataList.ToArray();
        
        EditorUtility.SetDirty(targetData);
        AssetDatabase.SaveAssets();
        
        return true;
    }
}
