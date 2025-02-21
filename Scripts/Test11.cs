using UnityEngine;
using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using ExcelDataReader;
public class Test11 : MonoBehaviour
{
    // 테스트용 GameData, 여기로 결과가 들어옴
    public GameData gameData = new GameData();

    void Start()
    {
        // Application.dataPath 예: "C:/Project/MyUnity/Assets"
        // 상위 폴더: "C:/Project/MyUnity"
        string parentFolder = Directory.GetParent(Application.dataPath).FullName;
        // 그 상위 폴더에 DataSheet 폴더가 있다고 가정
        string dataSheetFolder = Path.Combine(parentFolder, "DataSheet");

        Debug.Log($"[ExcelLoaderTest] Loading Excel files from: {dataSheetFolder}");

        // ExcelLoader 호출 (container=gameData, folderPath=dataSheetFolder)
        ExcelLoader.LoadAllExcelFiles(gameData, dataSheetFolder);

        // 로딩 결과 확인
        Debug.Log($"[ExcelLoaderTest] UnitDataList Count: {gameData.UnitData.Count}");
        if(gameData.UnitData.Count>0)
        {
            foreach (var item in gameData.UnitData)
            {
                Debug.Log($"{item.Key} //{item.Value.id} // {item.Value.abc}// {item.Value.bonusDamage[1]}// {item.Value.bonusDamage[2]}");
            }
        }
    }


}
