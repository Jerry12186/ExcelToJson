using System;
using System.Collections.Generic;

namespace Excel2Json
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> paths = CreateJosnFile.GetFlieFullName();

            for (int i = 0; i < paths.Count; i++)
            {
                ExcelToJson e2j = new ExcelToJson(paths[i]);
                CreateJosnFile.CreateText(e2j.ToJson(), e2j.GetSheetNames()[0]);
            }
            Console.WriteLine("生成完毕！按任意键退出。");
            Console.ReadKey();
        }


    }
}