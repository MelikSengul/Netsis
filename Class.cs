using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Concurrent;
using System.Threading;
using NetOpenX50;

namespace NetAi
{

    public static class fUtils
    {
        public static string StrToBool(this string obj)
        {
            return (obj == "True" || obj == "true" || obj == "1") ? "1" : "0";
        }

        public static string ToStr(this object obj)
        {
            return obj?.ToString() ?? "";
        }

        public static int ToInt(this object obj)
        {
            return int.TryParse(obj.ToStr(), out int result) ? result : 0;
        }

        public static double ToDbl(this object obj)
        {
            return double.TryParse(obj.ToStr(), out double result) ? result : 0;
        }

        public static bool ToBool(this object obj)
        {
            return bool.TryParse(obj.ToStr(), out bool result) && result;
        }

        public static DateTime ToDate(this object obj)
        {
            return DateTime.TryParse(obj.ToStr(), out DateTime result) ? result : DateTime.MinValue;
        }

        public static string ToSqlDbl(this object val)
        {
            try
            {
                return val.ToStr().Replace(',', '.');
            }
            catch
            {
                return "0";
            }
        }

        public static string ToSqlTrh(this DateTime value)
        {
            return value.ToString("yyyy-MM-dd");
        }

        public static string ToSqlTrhUzn(this DateTime date)
        {
            return date.ToString("yyyy-MM-dd HH:mm:ss");
        }
    }

    public class IniFile
    {
        public string path;

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        [DllImport("kernel32.DLL", EntryPoint = "GetPrivateProfileStringW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, string lpReturnString, int nSize, string lpFilename);

        public IniFile(string INIPath)
        {
            path = INIPath;
        }
        public void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, this.path);
        }
        public string IniReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, this.path);
            return temp.ToString();
        }
        public string IniReadValueDef(string Section, string Key, string def_val)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, def_val, temp, 255, this.path);
            return temp.ToString();
        }
        public List<string> GetCategories(string iniFile)
        {
            string returnString = new string(' ', 65536);
            GetPrivateProfileString(null, null, null, returnString, 65536, iniFile);
            List<string> result = new List<string>(returnString.Split('\0'));
            result.RemoveRange(result.Count - 2, 2);
            return result;
        }
        public List<string> GetKeys(string category)
        {
            string returnString = new string(' ', 32768);
            GetPrivateProfileString(category, null, null, returnString, 32768, this.path);
            List<string> result = new List<string>(returnString.Split('\0'));
            result.RemoveRange(result.Count - 2, 2);
            return result;
        }
    }
}