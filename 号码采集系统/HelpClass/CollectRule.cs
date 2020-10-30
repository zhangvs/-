using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace 号码采集系统
{

    class Reg360
    {
        public static string url = @"https://www.so.com/s?q=";
        public static string div = "<ul class=\"result\"><li[\\s\\S]*</li></ul>";
        public static string title = @"<h3 class=\""res-title \""[\s\S]*?</h3>";
        public static string titleText = @"<a\s*[^>]*>([\s\S]+?)</a>";
        public static string titleLink = "<a href=\"[\\s\\S]*?\"";
        public static string remark = "\n<p class=\"res-desc\"[\\s\\S]*?...</p>";
        public static string error = "系统检测到您操作过于频繁";
    }

    class RegBaidu
    {
        public static string url = @"https://www.baidu.com/s?wd=";
        public static string div = "<div id=\"content_left\"[\\s\\S]*<div style=\"clear:both;height:0;\"></div>";
        public static string title = @"<h3 class=\""t\""[\s\S]*?</h3>";
        public static string titleText = @"<a\s*[^>]*>([\s\S]+?)</a>";
        public static string titleLink = @"http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?";
        public static string remark = "<div class=\"c-abstract\"[\\s\\S]*?...</div>";
        public static string error = "您的电脑或所在的局域网络有异常的访问";
    }

    class CollectRule
    {
        public static bool HasChinese(string str)
        {
            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]");
        }
        //山东省临沂市，匹配全国省市。可是不行
        public static bool isCity(string str)
        {
            return Regex.IsMatch(str, @".+\u7701.+\u5e02.*");
        }
        public static bool isNumberZi(string str)
        {
            return Regex.IsMatch(str, @"\d+室\D|(东|西|南|北|转|角|侧)\d+米|[a-z|A-Z]座|[a-z|A-Z]区");
        }
        public static bool isAddress(string str)
        {
            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]市[\u4e00-\u9fa5]区|[\u4e00-\u9fa5]市[\u4e00-\u9fa5]县");
        }
        public static bool isLu(string str)
        {
            return Regex.IsMatch(str, @"[\s\S]+路(东|西|南|北)|(路|街|道)\d+(号|巷)");
        }
        public static bool isZi(string str)
        {
            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]+(广场|大街|大厦|小区|(软件|工业|软件|科技)园|交汇|有限公司|有限责任公司|地址：|地址:)");
        }

        //不包含“号段”和“靓号”的
        public static bool noIndex(string content)
        {
            if (content.IndexOf("号段") < 0 && content.IndexOf("靓号") < 0 && content.IndexOf("永久地址") < 0 && content.IndexOf("三区九县") < 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool isIndexKeyword(string content,string keyword)
        {
            if (keyword.Substring(0, 1) == "0")
            {
                int k = keyword.IndexOf('-');
                if (k > 0)
                {
                    keyword = keyword.Substring(k + 1, 7);
                }
                else
                {
                    keyword = keyword.Substring(4, 7);
                }
            }

            string keyword2 = keyword + "的问题";
            if (content.IndexOf(keyword) > 0 && content.IndexOf(keyword2) < 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //判断采集规则是否符合
        public static bool ruleJudge(string abstractStr,string keyword)
        {
           if( HasChinese(abstractStr) &&
                    noIndex(abstractStr) && isIndexKeyword(abstractStr, keyword) &&
                    (isNumberZi(abstractStr) || isAddress(abstractStr) || isLu(abstractStr) || isZi(abstractStr)))
            {
                return true;
            }
            else
            {
                return false;
            }
            
        }


        public static void getSearch(string search, ref string url, ref string error, ref string div)
        {
            if (search == "360")
            {
                url = Reg360.url;
                error = Reg360.error;
                div = Reg360.div;
            }
            else
            {

                url = RegBaidu.url;
                error = RegBaidu.error;
                div = RegBaidu.div;
            }
        }

        public static void getReg(string search, ref string title, ref string titleText, ref string titleLink, ref string remark)
        {
            if (search == "360")
            {
                title = Reg360.title;
                titleText = Reg360.titleText;
                titleLink = Reg360.titleLink;
                remark = Reg360.remark;
            }
            else
            {

                title = RegBaidu.title;
                titleText = RegBaidu.titleText;
                titleLink = RegBaidu.titleLink;
                remark = RegBaidu.remark;
            }
        }

        public static int MainWhile(string search, string keyword)
        {
            string url = "";
            string error = "";
            string div = "";
            getSearch(search, ref url, ref error, ref div);

            Regex r = new Regex(div, RegexOptions.IgnoreCase);
            string htmlStr = StringHelp.GetHtml2(url, keyword);
            string input = r.Match(htmlStr).Value;
            int bl = 0;
            if (htmlStr.IndexOf(error) > 0)
            {
                Object locker = new Object();
                lock (locker)
                {
                    try
                    {
                        StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() + " *************************************当前IP开始被" + search + "屏蔽*************************************" + keyword + "\r\n");
                        do
                        {
                            StringHelp.getChangeIp();
                            htmlStr = StringHelp.GetHtml2(url, keyword);
                        } while (htmlStr.IndexOf(error) > 0);

                        input = r.Match(htmlStr).Value;
                        if (input != "")
                        {
                            MainProc(search, keyword, input);
                        }
                    }
                    catch (Exception ex)
                    {
                        StringHelp.Write(StringHelp.pathError, DateTime.Now.ToString() + ex + keyword + "当前IP开始被" + search + "屏蔽,未启动VPN！\r\n");
                        bl = -1;
                    }
                }
            }
            else if (input != "")
            {
                if (MainProc(search,keyword, input))
                {
                    bl = 1;
                }
            }
            return bl;
        }

        public static bool MainProc(string search,string keyword, string input)
        {
            bool use = false;

            string title = "";
            string titleText = "";
            string titleLink = "";
            string remark = "";
            getReg(search, ref title, ref titleText, ref titleLink, ref remark);

            //标题
            Regex r = new Regex(title, RegexOptions.IgnoreCase);
            MatchCollection matchCollection = r.Matches(input);
            //简介
            string abstractReg = remark;
            MatchCollection abstractCollection = Regex.Matches(input, abstractReg, RegexOptions.IgnoreCase);

            for (int i = 0; i < abstractCollection.Count; i++)
            {
                Match m = abstractCollection[i];
                string abstractStr = StringHelp.ReplaceHtmlTag(m.Value.Replace("'", ""));
                if (m.Success && ruleJudge(abstractStr, keyword))
                {
                    //标题
                    string textReg = titleText;
                    MatchCollection textMatchCollection = Regex.Matches(matchCollection[i].Value, textReg, RegexOptions.IgnoreCase);
                    Match match2 = textMatchCollection[0];
                    string titleStr = "";
                    if (match2.Success)
                    {
                        titleStr = StringHelp.ReplaceHtmlTag(match2.Result("$1"));
                    }
                    //标题链接
                    string LinkReg = titleLink;
                    MatchCollection linkMatchCollection = Regex.Matches(matchCollection[i].Value, LinkReg, RegexOptions.IgnoreCase);
                    Match match3 = linkMatchCollection[0];
                    string titleLinkStr = "";
                    if (match3.Success)
                    {
                        titleLinkStr = match3.Value.Replace("<a href=", "").Replace("\"", "");
                    }

                    #region 创建采集对象
                    TelCollection tel = new TelCollection();
                    tel.Search = search;
                    tel.Telphone = keyword;
                    tel.Title = titleStr;
                    tel.TitleLink = titleLinkStr;
                    tel.Abstract = abstractStr;
                    #endregion
                    StringHelp.WriteExcel(tel);
                    use = true;
                    break;
                }
            }
            return use;

        }


        
    }
}
