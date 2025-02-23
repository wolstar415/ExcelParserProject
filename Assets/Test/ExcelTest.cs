using System.IO;
using UnityEngine;
public class ExcelTest : MonoBehaviour
{
    public ExcelData excelData = new ExcelData();

    void Start()
    {
        string parentFolder = Directory.GetParent(Application.dataPath).FullName;
        string dataSheetFolder = Path.Combine(parentFolder, "ExcelData");
        ExcelLoader.LoadAllExcelFiles(excelData, dataSheetFolder);
    }
}
