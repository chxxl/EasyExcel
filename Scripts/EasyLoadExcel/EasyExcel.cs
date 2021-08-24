using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

public class EasyExcel : EasyExcelUtil
{
    private string _path;
    private FileStream _fileStream;
    private IWorkbook _workbook;
    private ISheet _sheet;

    public EasyExcel(string path)
    {
        _path = path;

        try
        {
            _fileStream = new FileStream(_path, FileMode.Open, FileAccess.Read);
            _workbook = new XSSFWorkbook(_fileStream);
        }
        catch (FileNotFoundException)
        {
            DebugLog(LogType.Error, "������ �������� �ʽ��ϴ�.");
        }
        catch (System.Exception e)
        {
            if(e.HResult == 32)
                DebugLog(LogType.Error, "���� ������ �����ֽ��ϴ�.");
        }     
    }

    public bool CheckData(int row, int column, int sheetNum = 0)
    {
        if (_fileStream == null || _workbook == null)
        {
            DebugLog(LogType.Error, "�������� �����Ͱ� �������� �ʽ��ϴ�.");
            return false;
        }

         _sheet = _workbook.GetSheetAt(sheetNum);

        string cellData = GetCell(row, column).StringCellValue;

        // Blank
        if (cellData == string.Empty)
        {
            DebugLog(LogType.Normal, "�ش� ���� ����ֽ��ϴ�.");
            return false;
        }

        DebugLog(LogType.Normal, "�� ������ : " + cellData);
        return true;
    }

    public string GetData(int row, int column,int sheetNum = 0)
    {
       _sheet = _workbook.GetSheetAt(sheetNum);

        string data = string.Empty;

         if (GetCell(row, column).CellType == CellType.Numeric)
            data = GetCell(row, column).NumericCellValue.ToString();
        else data = GetCell(row, column).StringCellValue;

        return data;
    }

    public void SetData(int row, int column,string data, int sheetNum = 0)
    {
       _sheet = _workbook.GetSheetAt(sheetNum);

        GetCell(row, column).SetCellValue(data);    
    }

    public List<string> GetRowAllData(int row,int startColumn = 0, int sheetNum = 0)
    {
        _sheet = _workbook.GetSheetAt(sheetNum);

        IRow iRow = GetRow(row);
       List<string> rowDataArray = new List<string>(iRow.LastCellNum);

        for (int i = startColumn; i < iRow.LastCellNum; i++)
            rowDataArray.Add(iRow.GetCell(i).StringCellValue);

        return rowDataArray;
    }

    public List<string> GetColumnAllData(int column, int startRow = 0, int sheetNum = 0)
    {
        _sheet = _workbook.GetSheetAt(sheetNum);

        IRow irow = null;
        CellType type;
        List<string> dataList = new List<string>();

        for (int i = startRow;  ; i++)
        {
            irow = GetRow(i);

            type = irow.GetCell(column) != null ? irow.GetCell(column).CellType : CellType.Unknown;

            if (type == CellType.String)
            {
                if (irow.GetCell(column).StringCellValue == string.Empty)
                    break;

                dataList.Add(irow.GetCell(column).StringCellValue);
            }
            else if (type == CellType.Numeric)
            {
                dataList.Add(irow.GetCell(column).NumericCellValue.ToString());
            }
            else if (type == CellType.Unknown)
                break;
              
        }

        return dataList;
    }

    private ICell GetCell(int row, int column)
    {
        IRow rowData = GetRow(row);

        if (rowData.GetCell(column) == null)
            return rowData.CreateCell(column);

        return rowData.GetCell(column);
    }

    private IRow GetRow(int row)
    {
        if (_sheet.GetRow(row) == null)
            return _sheet.CreateRow(row);

        return _sheet.GetRow(row);
    }

    public void SaveFile()
    {
        try
        {
            FileStream file = new FileStream(_path, FileMode.Create, FileAccess.Write);
            _workbook.Write(file);

        }
        catch (System.Exception e)
        {
            Debug.Log(e.Message);
            DebugLog(LogType.Error, "���� ���� ����");
        }
    }



    // �׳� ȣ��� ���� �н� ���
    // �Ű������� �н� ���� �� �н��� ���

}
