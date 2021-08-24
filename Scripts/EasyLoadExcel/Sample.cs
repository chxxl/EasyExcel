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

        // ���ϴ� ���� �����Ͱ� �ִ��� üũ
        easyExcel.CheckData(0, 0);

        // ���ϴ� ���� �����͸� ������
        returnValue = easyExcel.GetData(1, 1);
        print(returnValue);

        // row ��� ������ ���
        returnList = easyExcel.GetRowAllData(0);
        PrintList(returnList);

        // Column ��� ������ ���
        returnList = easyExcel.GetColumnAllData(0);
        PrintList(returnList);

        // ���ϴ� ���� ������ ����
        easyExcel.SetData(2, 1, "Test");
        easyExcel.SaveFile();
    }

    private void PrintList(List<string> listData)
    {
        for (int i = 0; i < listData.Count; i++)        
            print(listData[i]);        
    }
}
