﻿using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;


public class XlsxReader
{
    [DllImport("kernel32.dll")]
    private static extern IntPtr _lopen(string lpPathName, int iReadWrite);
    [DllImport("kernel32.dll")]
    private static extern bool CloseHandle(IntPtr hObject);
    private const int OF_READWRITE = 2;
    private const int OF_SHARE_DENY_NONE = 0x40;
    private static readonly IntPtr HFILE_ERROR = new IntPtr(-1);

    private const string EXCEL_DATA_SHEET_NAME = "data$";
    private const string EXCEL_CONFIG_SHEET_NAME = "config$";
    private const int DATA_FIELD_DATA_START_INDEX = 5;
    /// <summary>
    /// 将指定Excel文件的内容读取到DataSet中
    /// </summary>
    public  DataSet ReadXlsxFile(string filePath, out string errorString)
    {
        // 检查文件是否存在且没被打开
        FileState fileState = GetFileState(filePath);
        if (fileState == FileState.Inexist)
        {
            errorString = string.Format("{0}文件不存在", filePath);
            return null;
        }
        else if (fileState == FileState.IsOpen)
        {
            errorString = string.Format("{0}文件正在被其他软件打开，请关闭后重新运行本工具", filePath);
            return null;
        }

        OleDbConnection conn = null;
        OleDbDataAdapter da = null;
        DataSet ds = null;

        try
        {
            // 初始化连接并打开
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\"";

            conn = new OleDbConnection(connectionString);
            conn.Open();

            // 获取数据源的表定义元数据                       
            DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

            // 必须存在数据表
            bool isFoundDateSheet = false;
            // 可选配置表
            bool isFoundConfigSheet = false;

            for (int i = 0; i < dtSheet.Rows.Count; ++i)
            {
                string sheetName = dtSheet.Rows[i]["TABLE_NAME"].ToString();

                if (sheetName == EXCEL_DATA_SHEET_NAME)
                    isFoundDateSheet = true;
                else if (sheetName == EXCEL_CONFIG_SHEET_NAME)
                    isFoundConfigSheet = true;
            }
            if (!isFoundDateSheet)
            {
                errorString = string.Format("错误：{0}中不含有Sheet名为{1}的数据表", filePath, EXCEL_DATA_SHEET_NAME.Replace("$", ""));
                return null;
            }

            // 初始化适配器
            da = new OleDbDataAdapter();
            da.SelectCommand = new OleDbCommand(String.Format("Select * FROM [{0}]", EXCEL_DATA_SHEET_NAME), conn);

            ds = new DataSet();
            da.Fill(ds, EXCEL_DATA_SHEET_NAME);

            // 删除表格末尾的空行
            DataRowCollection rows = ds.Tables[EXCEL_DATA_SHEET_NAME].Rows;
            int rowCount = rows.Count;
            for (int i = rowCount - 1; i >= DATA_FIELD_DATA_START_INDEX; --i)
            {
                if (string.IsNullOrEmpty(rows[i][0].ToString()))
                    rows.RemoveAt(i);
                else
                    break;
            }

            if (isFoundConfigSheet == true)
            {
                da.Dispose();
                da = new OleDbDataAdapter();
                da.SelectCommand = new OleDbCommand(String.Format("Select * FROM [{0}]", EXCEL_CONFIG_SHEET_NAME), conn);
                da.Fill(ds, EXCEL_CONFIG_SHEET_NAME);
            }
        }
        catch
        {
            errorString = "错误：连接Excel失败，你可能尚未安装Office数据连接组件: http://www.microsoft.com/en-US/download/details.aspx?id=23734 \n";
            return null;
        }
        finally
        {
            // 关闭连接
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
                // 由于C#运行机制，即便因为表格中没有Sheet名为data的工作簿而return null，也会继续执行finally，而此时da为空，故需要进行判断处理
                if (da != null)
                    da.Dispose();
                conn.Dispose();
            }
        }

        errorString = null;
        return ds;
    }
    /// <summary>
    /// 获取某个文件的状态
    /// </summary>
    private static FileState GetFileState(string filePath)
    {
        if (File.Exists(filePath))
        {
            IntPtr vHandle = _lopen(filePath, OF_READWRITE | OF_SHARE_DENY_NONE);
            if (vHandle == HFILE_ERROR)
                return FileState.IsOpen;

            CloseHandle(vHandle);
            return FileState.Available;
        }
        else
            return FileState.Inexist;
    }
    private enum FileState
    {
        Inexist,     // 不存在
        IsOpen,      // 已被打开
        Available,   // 当前可用
    }
}