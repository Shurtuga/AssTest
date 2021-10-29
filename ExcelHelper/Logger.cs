using CustomTools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomTools
{
    public enum LogMode
    {
        Regular,
        Urgent
    }
    public static class Logger
    {
        public static void Log(string message, string foldername, string filename, LogMode logMode, bool usespecialfolder)
        {
            string path;
            if (usespecialfolder) path = $"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}/{foldername}/{filename}";
            else path = filename;

            string addmes = "";
            if (logMode == LogMode.Urgent) addmes = "-- APP-BREAKING BUG ";

            using (StreamWriter sw = new StreamWriter(path, true))
            {
                sw.WriteLine($"{addmes}{DateTime.Now} : {message}");
            }
        }
        public static void Log(string message)
        {
            Log(message, "default", "log.txt", LogMode.Regular, false);
        }
        public static void Log(string message, string foldername)
        {
            Log(message, foldername, "log.txt", LogMode.Regular, false);
        }
        public static void Log(string message, string foldername, string filename)
        {
            Log(message, foldername, filename, LogMode.Regular, false);
        }
        public static void Log(string message, LogMode logMode)
        {
            Log(message, "default", "log.txt", logMode, false);
        }
        public static void Log(string message, bool usespecialfolder)
        {
            Log(message, "default", "log.txt", LogMode.Regular, usespecialfolder);
        }
        public static void Log(Exception e)
        {
            Log(GetFootprints(e), "default", "log.txt", LogMode.Urgent, false);
        }
        public static void Log(Exception e, string foldername, string filename, LogMode logMode, bool usespecialfolder)
        {
            string mes = GetFootprints(e);
            if (string.IsNullOrEmpty(mes)) mes = e.Message;
            Log(mes, foldername, filename, logMode, usespecialfolder);
        }
        static string GetFootprints(Exception e)
        {
            StackTrace st = new StackTrace(e, true);
            var frames = st.GetFrames();
            StringBuilder res = new StringBuilder();

            foreach (var frame in frames)
            {
                if (frame.GetFileLineNumber() < 1) continue;

                res.Append($"File: {frame.GetFileName()}");
                res.Append($", Method: {frame.GetMethod().Name}");
                res.Append($", Line: {frame.GetFileLineNumber()}");
                res.Append($" --> ");
            }
            return res.ToString();
        }
    }
}
