using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class EasyExcelUtil
{
   public enum LogType
    {
        Normal,
        Warning,
        Error
    }

    public void DebugLog(LogType type, string text)
    {
        switch (type)
        {
            case LogType.Normal:
                Debug.Log(text);
                break;
            case LogType.Warning:
                Debug.LogWarning(text);
                break;
            case LogType.Error:
                Debug.LogError(text);
                break;
            default:
                break;
        }
    }
}
