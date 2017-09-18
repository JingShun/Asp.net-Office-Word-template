using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

/// <summary>
/// Class_myDocx 的摘要描述
/// </summary>
public static class Class_myDocx
{
    public static string msg = "";
    /// <summary>
    /// 替換docx文檔中的解析器標籤(改用regex正則表示式)
    /// </summary>
    /// <param name="oxp">OpenXmlPart object</param>
    /// <param name="dct">字典包含要替換的解析器標籤</param>
    private static void parseRegex(OpenXmlPart oxp, Dictionary<string, string> dct, string 前綴詞 = "[$", string 後綴詞 = "$]")
    {
        //綴詞補上Regex的脫逸字元
        Regex 綴詞Reg = new Regex("(.)");
        前綴詞 = 綴詞Reg.Replace(前綴詞, ".*?\\$1");
        後綴詞 = 綴詞Reg.Replace(後綴詞, ".*?\\$1");

        //用StringBuilder來降低合併時效能問題
        StringBuilder 被取代字串 = new StringBuilder();
        string xmlString = null;
        Regex 正則取代=null;
        using (StreamReader sr = new StreamReader(oxp.GetStream()))
        { xmlString = sr.ReadToEnd(); }
        foreach (string key in dct.Keys)
        {
            //處理正則字串(篩選被取代字串，範例： \[\$.*?姓名.*?\$\])
            被取代字串.Clear(); //清空之前的內容
            //綴詞Reg.Replace(key, ".*?$1")：被取代字串的每個字元插入Regex(盡可能少的任意字符)
            被取代字串.Append(前綴詞).Append(綴詞Reg.Replace(key, ".*?$1")).Append(後綴詞);//合併要取代的Regex字串

            msg += "\r\n"+被取代字串;

            //開始取代
            正則取代 = new Regex(被取代字串.ToString());
            xmlString = 正則取代.Replace(xmlString, dct[key]);
        }

        using (StreamWriter sw = new StreamWriter(oxp.GetStream(FileMode.Create)))
        { sw.Write(xmlString); }
    }


    /// <summary>
    /// 替換docx文檔中的解析器標籤
    /// </summary>
    /// <param name="oxp">OpenXmlPart object</param>
    /// <param name="dct">字典包含要替換的解析器標籤</param>
    private static void parse(OpenXmlPart oxp, Dictionary<string, string> dct, string 前綴詞, string 後綴詞)
    {
        //用StringBuilder來降低合併時效能問題
        StringBuilder 被取代字串 = new StringBuilder();
        StringBuilder xmlString = new StringBuilder();
        using (StreamReader sr = new StreamReader(oxp.GetStream()))
        { xmlString.Append(sr.ReadToEnd()); }
        foreach (string key in dct.Keys)
        {
            被取代字串.Clear(); //清空之前的內容
            被取代字串.Append(前綴詞).Append(key).Append(後綴詞);//合併要取代的字串
            xmlString = xmlString.Replace(被取代字串.ToString(), dct[key]);
        }

        using (StreamWriter sw = new StreamWriter(oxp.GetStream(FileMode.Create)))
        { sw.Write(xmlString); }
    }

    /// <summary>
    ///解析模板文件並替換所有解析器標籤，返回二進制內容的新 docx 檔案.
    /// </summary>
    /// <param name="templateFile">模板文件路徑</param>
    /// <param name="dct">一個包含解析器標籤和值的字典</param>
    /// <param name="前綴詞">預設 [$ </param>
    /// <param name="後綴詞">預設 $] </param>
    /// <param name="Regex">預設 false，是否啟用正則表示式來處理標籤。若啟用，犧牲點效能來處理字串間多出的微軟標籤 </param>
    /// <returns></returns>
    public static byte[] MakeDocx(string templateFile, Dictionary<string, string> dct,string 前綴詞= "[$", string 後綴詞= "$]",bool Regex =false)
    {
        string tempFile = Path.GetTempPath() + ".docx";
        if(File.Exists(tempFile))
            File.Delete(tempFile);
        File.Copy(templateFile, tempFile);
        
        using (WordprocessingDocument wd = WordprocessingDocument.Open(
            tempFile, true))
        {
            Debug.WriteLine("body start");
            //Replace document body
            if (!Regex) parse(wd.MainDocumentPart, dct, 前綴詞, 後綴詞);
            else parseRegex(wd.MainDocumentPart, dct, 前綴詞, 後綴詞);
            
            Debug.WriteLine("Header start");
            foreach (HeaderPart hp in wd.MainDocumentPart.HeaderParts)
                if (!Regex) parse(hp, dct, 前綴詞, 後綴詞);
                else parseRegex(hp, dct, 前綴詞, 後綴詞);

            Debug.WriteLine("Footer start");
            foreach (FooterPart fp in wd.MainDocumentPart.FooterParts)
                if (!Regex) parse(fp, dct, 前綴詞, 後綴詞);
                else parseRegex(fp, dct, 前綴詞, 後綴詞);
        }
        byte[] buff = File.ReadAllBytes(tempFile);
        File.Delete(tempFile);
        return buff;
    }
}