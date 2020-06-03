using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Utils
{
    
    public static class Tools
    {
        public const string NumberSepr = ",";
        public const int MaxWhileCount = 10000;

        public static string HRMDateTime(DateTime d)
        {
            return string.Format("{0:dd/MM/yyyy HH:mm}", d);
        }
        public static string HRMDate(DateTime d)
        {
            return string.Format("{0:dd/MM/yyyy}", d);
        }
        private static string GetLanguageText_(string KEY, dynamic d)
        {
            string r = "";
            try
            {
                r = d[KEY].ToString();
            }
            catch { r = KEY; }
            return r;
        }
        public static string ReplaceStringLangValue_(string s, dynamic d)
        {
            int i; int j;
            i = s.IndexOf("{");
            j = s.IndexOf("}");
            if (i == -1 || j == -1 || i == j)
                return s;
            else
            {
                int iWhile = 0; // chặn lỗi Out Of Memory
                while (!(i == -1 || j == -1 || i == j) && iWhile < MaxWhileCount)
                {
                    iWhile++;
                    string s1 = s.Substring(i + 1, j - i - 1);
                    s = s.Replace("{" + s1 + "}", GetLanguageText_(s1, d));
                    i = s.IndexOf("{");
                    j = s.IndexOf("}");
                }
                return s;
            }
        }
        public static string insertSepr(string d0)
        {
            var d = "" + d0; // convert to string
            var i = 0;
            var d2 = "";
            var ic = 0;
            var ofs = d.Length - 1;
            var decimalpoint = d.IndexOf('.');
            if (decimalpoint >= 0) ofs = decimalpoint - 1;
            for (i = ofs; i >= 0; i--)
            {
                if (d[i].ToString() != NumberSepr)
                {
                    if (ic++ % 3 == 0 && i != ofs && d[i] != '-') d2 = NumberSepr + d2;
                    d2 = d[i] + d2;
                }
            }

            if (decimalpoint >= 0)
            {
                for (i = decimalpoint; i < d.Length; i++)
                    d2 += d[i];
            }
            return d2;
        }
        public static string FormatNumber(string num, int iRound = 0)
        {
            double a = 0;
            try
            {
                a = double.Parse(num);
                a = Math.Round(a, iRound);
            }
            catch
            {
                a = 0;
            }
            //if (a == 0) return "";
            return insertSepr(a.ToString());
            /*
            if (!IsHRMNumber(num)) num = "0";
            if (num == "0") return "";
            if (IsRound)
                return string.Format("{0:n0}", Double.Parse(num));
            else
                return string.Format("{0:n}", Double.Parse(num));
                */
        }
        private static object DataSetGet(string DataType, object val, int IsFormat = 0, dynamic d = null)
        {
            object r;
            string v = "";
            if (val != null) v = val.ToString();
            if (v.ToLower() == "null") v = "";
            v = v.Trim().Replace("{}", "");
            try
            {
                switch (DataType.ToLower())
                {
                    case "smallint": // string integer
                    case "int": // string integer
                    case "int32": // string integer
                    case "int16": // string integer
                        if (IsFormat == 1)
                            r = FormatNumber(v, 0);
                        else if (IsFormat == 2)
                        {
                            if (val != null) r = int.Parse(v); else r = 0;
                        }
                        else
                        {
                            if (val != null) r = val; else r = 0;
                        }
                        break;
                    case "int64": // string long
                    case "bigint": // string long
                        if (IsFormat == 1)
                            r = FormatNumber(v, 0);
                        else if (IsFormat == 2)
                        {
                            if (val != null) r = long.Parse(v); else r = 0;
                        }
                        else
                        {
                            if (val != null) r = val; else r = 0;
                        }
                        break;
                    case "double": // string double
                    case "decimal": // string double
                    case "numeric": // string double
                        if (IsFormat == 1)
                            r = FormatNumber(v, 0);
                        else if (IsFormat == 2)
                        {
                            if (val != null) r = double.Parse(v); else r = 0;
                        }
                        else
                        {
                            if (val != null) r = val; else r = 0;
                        }
                        break;
                    case "date": // string date
                        if (IsFormat == 1)
                            r = HRMDate(DateTime.Parse(v));
                        else if (IsFormat == 2)
                            r = DateTime.Parse(v);
                        else
                            r = val;
                        break;
                    case "datetime": // string date
                        if (IsFormat == 1)
                            r = HRMDateTime(DateTime.Parse(v));
                        else if(IsFormat == 2)
                            r = DateTime.Parse(v);
                        else
                            r = val;
                        break;
                    default:
                        if (d != null)
                        {
                            r = Tools.ReplaceStringLangValue_(v, d);
                        }
                        else
                        {
                            r = v;
                        }
                        break;
                }
            }
            catch (Exception)
            {
                switch (DataType.ToLower())
                {
                    case "smallint": // string integer
                    case "int": // string integer
                    case "int32": // string integer
                    case "int16": // string integer
                    case "bigint": // string long
                    case "int64": // string long
                    case "double": // string double
                    case "decimal": // string double
                    case "numeric": // string double
                        if (IsFormat == 1)
                            r = "";
                        else
                            r = 0;
                        break;
                    case "date": // string date
                    case "datetime": // string date
                    default:
                        r = "";
                        break;
                }
            }
            return r;
        }
        public static object GetDataJson(dynamic dL, dynamic d, int i, string sKey, int j = 0, int IsFormat = 0)
        {
            sKey = sKey.ToUpper();
            try
            {
                var itemType = (j == 0 ? "ItemType" : "ItemType" + j);
                var items = (j == 0 ? "Items" : "Items" + j);
                dynamic token = d[items][i][sKey];
                bool vIsNull = (token == null) ||
                        (token.Type == JTokenType.Array && !token.HasValues) ||
                        (token.Type == JTokenType.Object && !token.HasValues) ||
                        (token.Type == JTokenType.String && token.ToString() == String.Empty) ||
                        (token.Type == JTokenType.Null);

                return DataSetGet(d[itemType][sKey].ToString(), (vIsNull? null: token), IsFormat, dL);
            }
            catch (Exception)
            {
                return "0";
            }
        }
        
        public static object GetDataJson(dynamic d, string sKey, string DataType = "string")
        {
            sKey = sKey.ToUpper();
            return DataSetGet(DataType, d[sKey], 0);
        }
    }
}