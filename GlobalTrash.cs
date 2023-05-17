using System;
using System.Collections.Generic;
using System.Text;
namespace SWIFT_СПФС
{
    class GlobalTrash
    {
        public static string appPath = System.IO.Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
        //Создание объекта, для работы с файлом
        public static INIManager manager = new INIManager(appPath + "/SPFS.ini");

        //Получить значение по ключу
        public static string CodeBnk = manager.GetPrivateString("General", "CodeBnk");
        public static string SWIFTBnk = manager.GetPrivateString("General", "SWIFTBnk");
        public static string NameBnk = manager.GetPrivateString("General", "NameBnk");
        public static string Mist = manager.GetPrivateString("General", "Mist");
        public static string PathABSOut = manager.GetPrivateString("SWIFT", "PathABSOut");
        public static string PathABSIn = manager.GetPrivateString("SWIFT", "PathABSIn");
        public static string FileMaskOut = manager.GetPrivateString("SWIFT", "FileMaskOut");
        public static string PathTransit = manager.GetPrivateString("SWIFT", "PathTransit");
        public static string PathOut = manager.GetPrivateString("SWIFT", "PathOut");
        public static string PathIn = manager.GetPrivateString("SWIFT", "PathIn");
        public static string FileMaskIn = manager.GetPrivateString("SWIFT", "FileMaskIn");
        public static string PathArc = manager.GetPrivateString("SWIFT", "PathArc");
        public static string SPFSPathIn = manager.GetPrivateString("SPFS", "SPFSPathIn");
        public static string FileMask = manager.GetPrivateString("SPFS", "FileMask");
        public static string SPFSPathOut = manager.GetPrivateString("SPFS", "SPFSPathOut");
        public static string SPFSPathTransit = manager.GetPrivateString("SPFS", "SPFSPathTransit");
        public static string SPFSPathArc = manager.GetPrivateString("SPFS", "SPFSPathArc");

        public static DateTime datePickerFrom = new DateTime(2021, 2, 10);
        public static DateTime datePickerTo = DateTime.Now.AddDays(1);
    }
}