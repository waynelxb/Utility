using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeployFilesToTargetFolder
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceFileName = args[0];
            string sourcePath = args[1];
            string targetPath = args[2];

            //string sourcePath = @"C:\Projects\Finance_Repository\Code\Finance\Finance_Python";
            //string targetPath = @"C:\Projects\Finance_Repository";

            int i = 0;

            foreach (string newPath in Directory.GetFiles(sourcePath, sourceFileName, SearchOption.TopDirectoryOnly))
            {
                File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);

                Console.WriteLine("Deployed " + newPath + " as " + newPath.Replace(sourcePath, targetPath));
                i++;
            }

            Console.WriteLine("Summary: Deployed " + i.ToString() + " " + sourceFileName + " file(s).");

        }
    }
}
