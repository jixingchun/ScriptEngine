using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace ScriptEngine
{
    #region 工具类
    public class TypeUtils
    {
        public static object DispatchInvokeMember(string strMatchName, BindingFlags bindingFlags, Binder binder, object targetObj, object[] argments)
        {
            object obj = null;

            try
            {
                obj = targetObj.GetType().InvokeMember(strMatchName, bindingFlags, binder, targetObj, argments);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"DispatchInvokeMember : strMatchName={strMatchName} Exception={ex}");
                throw;
            }

            return obj;
        }
    }

    internal static class ConvertChineseChar
    {
        public static Dictionary<string, string> dicChinaToEn = new Dictionary<string, string>();
        public static Dictionary<string, string> dicEnToChina = new Dictionary<string, string>();
        public static List<string> EnEqualChinaLists = new List<string>();

        public static bool IsChina(string source, ref string destination)
        {
            if (string.IsNullOrEmpty(source))
            {
                return false;
            }
            if (EnEqualChinaLists.Contains(source))
            {
                destination = source;
                return true;
            }
            if (dicChinaToEn.ContainsKey(source))
            {
                destination = dicChinaToEn[source];
                return true;
            }
            int ii = 0;
            int jj = 0;
            int kk = source.Length;
            bool strSkip = false;
            bool rowSkip = false;
            string strTmp = "";
            bool boolValue = false;
            for (ii = 0, jj = 0; ii < kk; ii++, jj++)
            {
                if (source[ii] == '\'')
                {
                    rowSkip = true;
                    strSkip = true;
                }
                else
                {
                    if (!rowSkip)
                    {
                        if (source[ii] == '"')
                        {
                            strSkip = !strSkip;
                            continue;
                        }
                    }
                }

                if (strSkip)
                {
                    if ((rowSkip) && (source[ii] == '\n'))
                    {
                        strSkip = false;
                        rowSkip = false;
                    }
                    continue;
                }
                if ((source[ii] > 0x43FF) && (source[ii] < 0x9FA6))
                {
                    jj += 7;
                }
            }

            strSkip = false;
            rowSkip = false;

            char[] strDesTmp = new char[jj];
            for (ii = 0, jj = 0; ii < kk; ii++, jj++)
            {
                strDesTmp[jj] = source[ii];
                if (source[ii] == '\'')
                {
                    rowSkip = true;
                    strSkip = true;
                }
                else
                {
                    if (!rowSkip)
                    {

                        if (source[ii] == '"')
                        {
                            strSkip = !strSkip;
                            continue;
                        }
                    }
                }

                if (strSkip)
                {
                    if ((rowSkip) && (source[ii] == '\n'))
                    {
                        strSkip = false;
                        rowSkip = false;

                    }

                    continue;
                }
                if ((source[ii] > 0x43FF) && (source[ii] < 0x9FA6))
                {
                    boolValue = true;
                    strTmp = ((UInt16)(source[ii])).ToString("x4");
                    strDesTmp[jj++] = 'U';
                    strDesTmp[jj++] = '_';
                    strDesTmp[jj++] = 'U';
                    strDesTmp[jj++] = '_';
                    strDesTmp[jj++] = strTmp[0];
                    strDesTmp[jj++] = strTmp[1];
                    strDesTmp[jj++] = strTmp[2];
                    strDesTmp[jj] = strTmp[3];
                }
            }
            destination = new string(strDesTmp);
            if (!boolValue)
                EnEqualChinaLists.Add(source);
            else
            {
                if (!dicChinaToEn.ContainsKey(source))
                    dicChinaToEn.Add(source, destination);
                if (!dicEnToChina.ContainsKey(destination))
                    dicEnToChina.Add(destination, source);
            }
            strDesTmp = null;
            return boolValue;
        }

        public static string UnicodeToChinese(string strSource)
        {
            if (string.IsNullOrEmpty(strSource))
                return "";
            if (EnEqualChinaLists.Contains(strSource))
                return strSource;
            if (dicEnToChina.ContainsKey(strSource))
                return dicEnToChina[strSource];
            int ii = 0;
            int jj = 0;
            int kk = 0;
            bool strSkip = false;
            bool rowSkip = false;
            byte valTmp = 0;
            int valSum = 0;
            int changeCnt = 0;
            int matchCount = Regex.Matches(strSource, "U_U_").Count;

            for (ii = 0; ii < strSource.Length; ii++)
            {

                if (rowSkip)
                {
                    if ((strSource[ii] == 'U') && (strSource[ii + 1] == '_') && (strSource[ii + 2] == 'U') &&
                        (strSource[ii + 3] == '_'))
                    {
                        ii += 7;
                        kk++;
                    }

                    if (strSource[ii] == '\n')
                    {
                        rowSkip = false;
                    }

                }

            }

            bool isConvert = false;
            jj = strSource.Length - ((matchCount - kk) * 7);
            char[] strDesTmp = new char[jj];
            for (ii = 0, jj = 0; jj < strDesTmp.Length; ii++, jj++)
            {
                strDesTmp[jj] = strSource[ii];
                if (strSource[ii] == '\'')
                {
                    //rowSkip = true;
                    //strSkip = true;
                }
                else
                {
                    if (!rowSkip)
                    {
                        if (strSource[ii] == '"')
                        {
                            strSkip = !strSkip;
                            continue;
                        }
                    }
                }
                if (strSkip)
                {
                    if ((rowSkip) && (strDesTmp[jj] == '\n'))
                    {
                        strSkip = false;
                        rowSkip = false;
                    }
                    continue;
                }

                if (changeCnt < matchCount)
                {
                    if ((strSource[ii] == 'U') && (strSource[ii + 1] == '_') && (strSource[ii + 2] == 'U') &&
                        (strSource[ii + 3] == '_'))
                    {
                        isConvert = true;
                        valTmp = Convert.ToByte(strSource[ii + 4].ToString(), 16);
                        valSum = valTmp * 4096;
                        valTmp = Convert.ToByte(strSource[ii + 5].ToString(), 16);
                        valSum += (valTmp * 256);
                        valTmp = Convert.ToByte(strSource[ii + 6].ToString(), 16);
                        valSum += (valTmp * 16);
                        valTmp = Convert.ToByte(strSource[ii + 7].ToString(), 16);
                        valSum += valTmp;
                        strDesTmp[jj] = (char)(valSum);
                        ii += 7;
                        changeCnt++;
                    }
                }
            }
            string strDes = new string(strDesTmp);
            strDesTmp = null;
            if (!isConvert)
            {
                EnEqualChinaLists.Add(strSource);
                return strSource;
            }
            else
            {
                if (!dicEnToChina.ContainsKey(strSource))
                    dicEnToChina.Add(strSource, strDes);
                if (!dicChinaToEn.ContainsKey(strDes))
                    dicChinaToEn.Add(strDes, strSource);
            }

            return strDes;

        }
    }
    #endregion
}
