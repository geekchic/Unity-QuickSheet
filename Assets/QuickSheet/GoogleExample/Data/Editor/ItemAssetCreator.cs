using UnityEngine;
using UnityEditor;
using System.IO;
using UnityQuickSheet;

///
/// !!! Machine generated code !!!
/// 
public partial class GoogleDataAssetUtility
{
    [MenuItem("Assets/Create/Google/Item")]
    public static void CreateItemAssetFile()
    {
        Item asset = CustomAssetUtility.CreateAsset<Item>();
        asset.SheetName = "1_OXrt1si7FVTZGyuXEgg8WVca6nYaU6GbwN1frcvZBc";
        asset.WorksheetName = "Item";
        EditorUtility.SetDirty(asset);        
    }
    
}