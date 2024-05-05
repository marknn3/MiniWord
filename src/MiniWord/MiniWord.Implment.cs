using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MiniSoftware.Extensions;
using MiniSoftware.Utility;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MiniSoftware;

public static partial class MiniWord
{
    private static void SaveAsByTemplateImpl(Stream stream, byte[] template, Dictionary<string, object> data)
    {
        var value = data;
        byte[] bytes = null;
        using (var ms = new MemoryStream())
        {
            ms.Write(template, 0, template.Length);
            ms.Position = 0;
            using (var docx = WordprocessingDocument.Open(ms, true))
            {
                var hc = docx.MainDocumentPart.HeaderParts.Count();
                var fc = docx.MainDocumentPart.FooterParts.Count();
                for (int i = 0; i < hc; i++)
                {
                    docx.MainDocumentPart.HeaderParts.ElementAt(i).Header.Generate(docx, value);
                }
                for (int i = 0; i < fc; i++)
                {
                    docx.MainDocumentPart.FooterParts.ElementAt(i).Footer.Generate(docx, value);
                }
                docx.MainDocumentPart.Document.Body.Generate(docx, value);
                docx.Save();
            }
            bytes = ms.ToArray();
        }
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void Generate(this OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
    {
        // avoid {{tag}} like <t>{</t><t>{</t> 
        // avoid {{tag}} like <t>aa{</t><t>{</t>  test in...
        AvoidSplitTagText(xmlElement);

        ReplaceTables(xmlElement, docx, tags);

        ReplaceStatements(xmlElement, tags);

        ReplaceText(xmlElement, docx, tags);
    }

    private static void ReplaceTables(OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
    {
        var tables = xmlElement.Descendants<Table>().ToArray();
        {
            foreach (var table in tables)
            {
                var trs = table.Descendants<TableRow>().ToArray(); // remember toarray or system will loop OOM;

                foreach (var tr in trs)
                {
                    var innerText = tr.InnerText
                        .Replace("{{foreach", "")
                        .Replace("endforeach}}", "")    
                        .Replace("{{if(", "")
                        .Replace("}}else{{", "")
                        .Replace(")if", "")
                        .Replace("endif}}", "");
                    var matches = Regex.Matches(innerText, @"(?<={{).+?\..+?(?=}})")
                        .Cast<Match>().GroupBy(x => x.Value)
                        .Select(varGroup => varGroup.First().Value).ToArray();
                    if (matches.Length > 0)
                    {
                        var listKeys = matches.Select(s => s.Substring(0, s.LastIndexOf('.'))).Distinct().ToArray();
                        // TODO:
                        // not support > 2 list in same tr
                        if (listKeys.Length > 2)
                        {
                            throw new NotSupportedException("MiniWord doesn't support more than 2 list in same row");
                        }
                        var listKey = listKeys[0];
                        if (tags.ContainsKey(listKey) && tags[listKey] is IEnumerable)
                        {
                            var list = tags[listKey] as IEnumerable;

                            foreach (Dictionary<string, object> es in list)
                            {
                                var dic = new Dictionary<string, object>(); //TODO: optimize

                                var newTr = tr.CloneNode(true);
                                foreach (var e in es)
                                {
                                    var dicKey = $"{listKey}.{e.Key}";
                                    dic.Add(dicKey, e.Value);
                                }

                                ReplaceStatements(newTr, tags: dic);
                                ReplaceText(newTr, docx, tags: dic);

                                //Fix #47 The table should be inserted at the template tag position instead of the last row
                                if (table.Contains(tr))
                                {
                                    table.InsertBefore(newTr, tr);
                                }
                                else
                                {
                                    // If it is a nested table, temporarily append it to the end according to the original plan.
                                    table.Append(newTr);
                                }
                            }
                            tr.Remove();
                        }
                    }
                }
            }
        }
    }

    private static void AvoidSplitTagText(OpenXmlElement xmlElement)
    {
        var texts = xmlElement.Descendants<Text>().ToList();
        var pool = new List<Text>();
        var sb = new StringBuilder();
        var needAppend = false;
        for (int i = 0; i < texts.Count; i++)
        {
            Text text = texts[i];
            var clear = false;
            if (!needAppend)
            {
                if (text.InnerText.StartsWith("{{"))
                {
                    needAppend = true;
                }
                else if ((i + 1 < texts.Count && text.InnerText.EndsWith("{") && texts[i + 1].InnerText.StartsWith("{"))
                      || (text.InnerText.Contains("{{")))
                {
                    int pos = text.InnerText.IndexOf("{");
                    if (pos > 0)
                    {
                        // Split text to always start with {{
                        var newText = text.Clone() as Text;
                        newText.Text = text.InnerText.Substring(0, pos);
                        text.Text = text.InnerText.Substring(pos);
                        text.Parent.InsertBefore(newChild: newText, text);
                    }
                    needAppend = true;
                }
            }
            if (needAppend)
            {
                sb.Append(text.InnerText);
                pool.Add(text);

                var s = sb.ToString();
                // TODO: check tag exist
                // TODO: record tag text if without tag then system need to clear them
                // TODO: every {{tag}} one <t>for them</t> and add text before first text and copy first one and remove {{, tagname, }}

                const string foreachTag = "{{foreach";
                const string endForeachTag = "endforeach}}";
                const string ifTag = "{{if";
                const string endifTag = "endif}}";
                const string tagStart = "{{";
                const string tagEnd = "}}";

                var foreachTagContains = s.Count(foreachTag) == s.Count(endForeachTag);
                var ifTagContains = s.Count(ifTag) == s.Count(endifTag);
                var tagContains = s.StartsWith(tagStart) && s.Contains(tagEnd);

                if (foreachTagContains && ifTagContains && tagContains)
                {
                    if (sb.Length <= 1000) // avoid too big tag
                    {
                        var first = pool.First();
                        var newText = first.Clone() as Text;
                        newText.Text = s;
                        first.Parent.InsertBefore(newText, first);
                        foreach (var t in pool)
                        {
                            t.Text = "";
                        }
                    }
                    clear = true;
                }
            }

            if (clear)
            {
                sb.Clear();
                pool.Clear();
                needAppend = false;
            }
        }
    }

    //private static void AvoidSplitTagText(OpenXmlElement xmlElement, IEnumerable<string> txt)
    //{
    //    foreach (var paragraph in xmlElement.Descendants<Paragraph>())
    //    {
    //        foreach (var continuousString in paragraph.GetContinuousString())
    //        {
    //            foreach (var text in txt.Where(o => continuousString.Item1.Contains(o)))
    //            {
    //                continuousString.Item3.TrimStringToInContinuousString(text);
    //            }
    //        }
    //    }
    //}

    //private static List<string> GetReplaceKeys(Dictionary<string, object> tags)
    //{
    //    var keys = new List<string>();
    //    foreach (var item in tags)
    //    {
    //        if (item.Value.IsStrongTypeEnumerable())
    //        {
    //            foreach (var item2 in (IEnumerable)item.Value)
    //            {
    //                if (item2 is Dictionary<string, object> dic)
    //                {
    //                    foreach (var item3 in dic.Keys)
    //                    {
    //                        keys.Add("{{" + item.Key + "." + item3 + "}}");
    //                    }
    //                }
    //                break;
    //            }
    //        }
    //        else
    //        {
    //            keys.Add("{{" + item.Key + "}}");
    //        }
    //    }
    //    return keys;
    //}

    private static bool EvaluateStatement(string tagValue, string comparisonOperator, string value)
    {
        var checkStatement = false;

        var tagValueEvaluation = EvaluateValue(tagValue);

        switch (tagValueEvaluation)
        {
            case double dtg when double.TryParse(value, out var doubleNumber):
                switch (comparisonOperator)
                {
                    case "==":
                    case "=":
                        checkStatement = dtg.Equals(doubleNumber);
                        break;
                    case "!=":
                    case "<>":
                        checkStatement = !dtg.Equals(doubleNumber);
                        break;
                    case ">":
                        checkStatement = dtg > doubleNumber;
                        break;
                    case "<":
                        checkStatement = dtg < doubleNumber;
                        break;
                    case ">=":
                        checkStatement = dtg >= doubleNumber;
                        break;
                    case "<=":
                        checkStatement = dtg <= doubleNumber;
                        break;
                }
                break;

            case int itg when int.TryParse(value, out var intNumber):
                switch (comparisonOperator)
                {
                    case "==":
                    case "=":
                        checkStatement = itg.Equals(intNumber);
                        break;
                    case "!=":
                    case "<>":
                        checkStatement = !itg.Equals(intNumber);
                        break;
                    case ">":
                        checkStatement = itg > intNumber;
                        break;
                    case "<":
                        checkStatement = itg < intNumber;
                        break;
                    case ">=":
                        checkStatement = itg >= intNumber;
                        break;
                    case "<=":
                        checkStatement = itg <= intNumber;
                        break;
                }
                break;

            case DateTime dttg when DateTime.TryParse(value, out var date):
                switch (comparisonOperator)
                {
                    case "==":
                    case "=":
                        checkStatement = dttg.Equals(date);
                        break;
                    case "!=":
                    case "<>":
                        checkStatement = !dttg.Equals(date);
                        break;
                    case ">":
                        checkStatement = dttg > date;
                        break;
                    case "<":
                        checkStatement = dttg < date;
                        break;
                    case ">=":
                        checkStatement = dttg >= date;
                        break;
                    case "<=":
                        checkStatement = dttg <= date;
                        break;
                }
                break;

            case string stg:
                switch (comparisonOperator)
                {
                    case "==":
                    case "=":
                        checkStatement = stg == value;
                        break;
                    case "!=":
                    case "<>":
                        checkStatement = stg != value;
                        break;
                }
                break;

            case bool btg when bool.TryParse(value, out var boolean):
                switch (comparisonOperator)
                {
                    case "==":
                    case "=":
                        checkStatement = btg != boolean;
                        break;
                    case "!=":
                    case "<>":
                        checkStatement = btg == boolean;
                        break;
                }
                break;
        }

        return checkStatement;
    }

    private static object EvaluateValue(string value)
    {
        if (double.TryParse(value, out var doubleNumber))
        {
            return doubleNumber;
        }
        else if (int.TryParse(value, out var intNumber))
        {
            return intNumber;
        }
        else if (DateTime.TryParse(value, out var date))
        {
            return date;
        }

        return value;
    }

    private static void ReplaceText(OpenXmlElement xmlElement, WordprocessingDocument docx, Dictionary<string, object> tags)
    {
        var paragraphs = xmlElement.Descendants<Paragraph>().ToArray();
        foreach (Paragraph p in paragraphs)
        {
            var runs = p.Descendants<Run>().ToArray();

            foreach (Run run in runs)
            {
                var texts = run.Descendants<Text>().ToArray();

                foreach (Text text in texts)
                {
                    foreach (var tag in tags)
                    {
                        var isMatch = text.Text.Contains("{{" + tag.Key + "}}");

                        if (!isMatch && tag.Value is List<MiniWordForeach> forTags)
                        {
                            if (forTags.Any(forTag => forTag.Value.Keys.Any(
                                dictKey =>
                                {
                                    var innerTag = "{{" + tag.Key + "." + dictKey + "}}";
                                    return text.Text.Contains(innerTag);
                                })))
                            {
                                isMatch = true;
                            }
                        }

                        if (isMatch)
                        {
                            //if (tag.Value is string[] || tag.Value is IList<string> || tag.Value is List<string>)
                            if (tag.Value is IList<string>)
                            {
                                var vs = tag.Value as IEnumerable;
                                var currentT = text;
                                var isFirst = true;
                                foreach (object v in vs)
                                {
                                    var newT = text.CloneNode(true) as Text;
                                    newT.Text = text.Text.Replace("{{" + tag.Key + "}}", v?.ToString());
                                    if (isFirst)
                                    {
                                        isFirst = false;
                                    }
                                    else
                                    {
                                        run.Append(new Break());
                                    }
                                    newT.Text = EvaluateIfStatement(newT.Text);
                                    run.Append(newT);
                                    currentT = newT;
                                }
                                text.Remove();
                            }

                            else if (tag.Value is List<MiniWordForeach> vs)
                            {
                                var currentT = text;
                                var generatedText = new Text();
                                currentT.Text = currentT.Text.Replace("{{foreach", "").Replace("endforeach}}", "");

                                var newTexts = new Dictionary<int, string>();
                                for (var i = 0; i < vs.Count; i++)
                                {
                                    var newT = text.CloneNode(true) as Text;

                                    foreach (var vv in vs[i].Value)
                                    {
                                        newT.Text = newT.Text.Replace("{{" + tag.Key + "." + vv.Key + "}}", vv.Value.ToString());
                                    }

                                    newT.Text = EvaluateIfStatement(newT.Text);

                                    if (!string.IsNullOrEmpty(newT.Text))
                                    {
                                        newTexts.Add(i, newT.Text);
                                    }
                                }

                                for (var i = 0; i < newTexts.Count; i++)
                                {
                                    var dict = newTexts.ElementAt(i);
                                    generatedText.Text += dict.Value;

                                    if (i != newTexts.Count - 1)
                                    {
                                        generatedText.Text += vs[dict.Key].Separator;
                                    }
                                }

                                run.Append(generatedText);
                                text.Remove();
                            }

                            else if (tag.Value is IMiniWordComponent value)
                            {
                                value.Execute(docx, run, value);
                                text.Remove();
                            }

                            else if (tag.Value is IList<IMiniWordComponentList> valueList)
                            {
                                var firstValue = valueList.FirstOrDefault();
                                if (firstValue != null)
                                {
                                    if (valueList.Any(value => value.GetType() != firstValue.GetType()))
                                    {
                                        throw new NotSupportedException("MiniWord doesn't support covarient lists");
                                    }
                                    //if (valueList.Any(value => !(value is IMiniWordComponent)))
                                    //{
                                    //    throw new NotSupportedException("MiniWord doesn't support covarient lists");
                                    //}
                                    var list = tag.Value as IList<IMiniWordComponent>;
                                    firstValue?.Execute(docx, run, list);
                                }
                                text.Remove();
                            }

                            else
                            {
                                string newText;
                                if (tag.Value is DateTime)
                                {
                                    newText = ((DateTime)tag.Value).ToString("yyyy-MM-dd HH:mm:ss");
                                }
                                else
                                {
                                    newText = tag.Value?.ToString();
                                }

                                text.Text = text.Text.Replace("{{" + tag.Key + "}}", newText);
                            }
                        }
                    } // foreach (var tag in tags)

                    text.Text = EvaluateIfStatement(text.Text);

                    // add breakline
                    {
                        var newText = text.Text;
                        var splits = Regex.Split(newText, "(?:<[a-zA-Z/].*?>|\n)");
                        var currentT = text;
                        var isFirst = true;
                        if (splits.Length > 1)
                        {
                            foreach (var v in splits)
                            {
                                var newT = text.CloneNode(true) as Text;
                                newT.Text = v?.ToString();
                                if (isFirst)
                                {
                                    isFirst = false;
                                }
                                else
                                {
                                    run.Append(new Break());
                                }
                                run.Append(newT);
                                currentT = newT;
                            }
                            text.Remove();
                        }
                    }
                } // foreach (Text text in texts)
            } // foreach (Run run in runs)
        } // foreach (Paragraph p in paragraphs)
    }

    private static void ReplaceStatements(OpenXmlElement xmlElement, Dictionary<string, object> tags)
    {
        var paragraphs = xmlElement.Descendants<Paragraph>().ToList();

        while (paragraphs.Any(s => s.InnerText.Contains("@if")))
        {
            int ifIndex = paragraphs.FindLastIndex(s => s.InnerText.Contains("@if"));
            int elseIndex = paragraphs.FindIndex(0, s => s.InnerText.Contains("@else"));
            int endifIndex = paragraphs.FindIndex(ifIndex, s => s.InnerText.Contains("@endif"));

            var statement = paragraphs[ifIndex].InnerText.Split(' ');

            var tagValue = tags[statement[1]] ?? "NULL";

            var checkStatement = statement.Length == 4 ? EvaluateStatement(tagValue.ToString(), statement[2], statement[3]) : !bool.Parse(tagValue.ToString());

            if (checkStatement == true)
            {
                // Else present
                if (elseIndex != -1 && elseIndex < endifIndex)
                {
                    paragraphs[elseIndex].Remove();
                    for (int i = elseIndex + 1; i <= endifIndex - 1; i++)
                    {
                        paragraphs[i].Remove();
                    }
                }
            }
            else
            {
                // Else present
                if (elseIndex != -1 && elseIndex < endifIndex)
                {
                    for (int i = ifIndex + 1; i <= elseIndex - 1; i++)
                    {
                        paragraphs[i].Remove();
                    }
                    paragraphs[elseIndex].Remove();
                }
                else
                {
                    for (int i = ifIndex + 1; i <= endifIndex - 1; i++)
                    {
                        paragraphs[i].Remove();
                    }
                }
            }

            paragraphs[ifIndex].Remove();
            paragraphs[endifIndex].Remove();

            paragraphs = xmlElement.Descendants<Paragraph>().ToList();
        }
    }

    private static string EvaluateIfStatement(string text)
    {
        const string ifStartTag = "{{if(";
        const string ifEndTag = ")if";
        const string elseTag = "}}else{{";
        const string endIfTag = "endif}}";

        while (text.Contains(ifStartTag))
        {
            int ifIndex = text.LastIndexOf(ifStartTag, StringComparison.Ordinal);
            int ifEndIndex = text.IndexOf(ifEndTag, ifIndex, StringComparison.Ordinal);

            string[] statement = text.Substring(ifIndex + ifStartTag.Length, ifEndIndex - (ifIndex + ifStartTag.Length)).Split(',');

            bool checkStatement = EvaluateStatement(statement[0], statement[1], statement[2]);

            if (checkStatement == true)
            {
                text = text.Remove(ifIndex, ifEndIndex - ifIndex + ifEndTag.Length);
                int elseIndex = text.IndexOf(elseTag, ifIndex, StringComparison.Ordinal);
                int endIfFinalIndex = text.IndexOf(endIfTag, ifIndex, StringComparison.Ordinal);
                if (elseIndex != -1 && elseIndex < endIfFinalIndex)
                {
                    text = text.Remove(elseIndex, endIfFinalIndex - elseIndex + endIfTag.Length);
                }
                else
                {
                    text = text.Remove(endIfFinalIndex, endIfTag.Length);
                }
            }
            else
            {
                int elseIndex = text.IndexOf(elseTag, ifEndIndex, StringComparison.Ordinal);
                int endIfFinalIndex = text.IndexOf(endIfTag, ifEndIndex, StringComparison.Ordinal);
                if (elseIndex != -1 && elseIndex < endIfFinalIndex)
                {
                    int count = elseIndex - ifIndex + elseTag.Length;
                    text = text.Remove(ifIndex, count);
                    text = text.Remove(endIfFinalIndex - count, endIfTag.Length);
                }
                else
                {
                    text = text.Remove(ifIndex, endIfFinalIndex - ifIndex + endIfTag.Length);
                }
            }
        }

        return text;
    }


    private static byte[] GetBytes(string path)
    {
        using (var st = Helpers.OpenSharedRead(path))
        using (var ms = new MemoryStream())
        {
            st.CopyTo(ms);
            return ms.ToArray();
        }
    }
}