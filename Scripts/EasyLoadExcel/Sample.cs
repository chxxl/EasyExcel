using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class Sample : MonoBehaviour
{
    void Start()
    {
        string returnValue;
        List<string> returnList;

        EasyExcel easyExcel = new EasyExcel(Application.dataPath + "\\Test.xlsx");

        // 원하는 셀에 데이터가 있는지 체크
        easyExcel.CheckData(0, 0);

        // 원하는 셀의 데이터를 가져옴
        returnValue = easyExcel.GetData(1, 1);
        print(returnValue);

        // row 모든 데이터 출력
        returnList = easyExcel.GetRowAllData(0);
        PrintList(returnList);

        // Column 모든 데이터 출력
        returnList = easyExcel.GetColumnAllData(0);
        PrintList(returnList);

        // 원하는 셀에 데이터 삽입
        easyExcel.SetData(2, 1, "Test");
        easyExcel.SaveFile();
    }

    private void PrintList(List<string> listData)
    {
        for (int i = 0; i < listData.Count; i++)        
            print(listData[i]);        
    }
}
