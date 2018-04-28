
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
namespace Excel2Json
{
    class CreateJosnFile
    {
        /// <summary>
        /// 获取Excel的全路径
        /// </summary>
        /// <returns></returns>
        public static List<string> GetFlieFullName()
        {
            List<string> paths = new List<string>();
            string strPath = Environment.CurrentDirectory + "\\Excel";
            DirectoryInfo theFolder = new DirectoryInfo(strPath);

            foreach (FileInfo file in theFolder.GetFiles())
            {
                if (Path.GetExtension(file.Name) == ".xls" || Path.GetExtension(file.Name) == ".xlsx")
                {
                    paths.Add(strPath + "\\" + file.Name);
                }
            }

            return paths;
        }

        public static void CreateText(string content, string fileName)
        {
            Directory.CreateDirectory("Json");
            string filePath = "Json\\" + fileName + ".json";
            FileStream fs = new FileStream(filePath, FileMode.Create);
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            fs.Write(bytes, 0, bytes.Length);
            
            //清空缓冲区、关闭流
            fs.Flush();
            fs.Close();
        }
    }
}
