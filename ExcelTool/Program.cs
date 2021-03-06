﻿using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    class Program
    {
        private const string txtPath = @"D:\data\txt";
        private const string txtPostFix = ".txt";
        private const int dataDistance = 5;
        private const int tabDistance = 8;

        static void Main(string[] args)
        {
            if (args != null)
            {
                if(args[0]== "ReadLine")
                {

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("---------若有修改以上xlsx文件,需通知服务端进行处理。若无修改，请忽略---------");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("按任意键继续");
                    Console.ReadLine();
                    Console.ResetColor(); //将控制台的前景色和背景色设为默认值
                }
                else
                {
                    string fileName = Path.GetFileNameWithoutExtension(args[0]);     //返回不带扩展名的文件名 
                    string savepath = args[1];
                    string savefile = args[2];
                    Program p = new Program();
                    XlsxReader xr = new XlsxReader();
                    string errorString;
                    DataSet dataset = xr.ReadXlsxFile(args[0], out errorString);
                    p.ExportToTxt(dataset, savepath + @"\" + fileName + savefile);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("服务端文件{0}{1}导出完毕,存放目录为：{2}", fileName, savefile, savepath);
                    Console.ResetColor(); //将控制台的前景色和背景色设为默认值
                }

            }

        }
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="PathStr">完整路径</param>
        public void ExportToTxt(DataSet ds, string PathStr)
        {
            StreamWriter writer = new StreamWriter(PathStr, false, new UTF8Encoding(false));

           
            int n = 0;
            writer.WriteLine("%%配置导表自动生成，请不要随意手动修改！！！");
            writer.WriteLine("%%配置导表自动生成，请不要随意手动修改！！！");
            writer.WriteLine("%%配置导表自动生成，请不要随意手动修改！！！");
            foreach (DataRow mDr in ds.Tables[0].Rows)
            {
                if(n>=5)
                {
                    string s = "-define(";

                    s =s+ mDr[0].ToString() +","+ mDr[1].ToString()+").    %%" + mDr[2].ToString();
                    writer.WriteLine(s);
      
                }
                else
                {
                    n++;
                }
               
            }
            writer.Flush();
            writer.Close();
        }
    }
}