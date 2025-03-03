using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;

namespace TarjimonWord
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;
            //var selection = wordApp;
            //selection.Text=KrillToLatin(selection.Text);
            var content = doc.Content;
            content.Text= KrillToLatin(content.Text);
        }
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = Globals.ThisAddIn.Application;
            var doc = wordApp.ActiveDocument;
            var content = doc.Content;
            content.Text = LotinToKrill(content.Text);
        }

        private string KrillToLatin(string text)
        {
            var _dict = new Dictionary<string, string> {
                { "Ў","O‘" }, { "ў","o‘"},{ "Ғ","G‘" }, { "ғ","g‘" },
                { "Ҳ","H"}, { "ҳ","h"}, { "Қ","Q" }, { "қ","q" }
            };
            var _result = string.Join("", text.Select(c => _dict.ContainsKey(c.ToString()) ? _dict[c.ToString()] : c.ToString()));

            var dict = new Dictionary<string, string>
            {
                
                { "А", "A" }, { "Б", "B" }, { "В", "V" }, { "Г", "G" }, { "Д", "D" },
                { "Е", "E" }, { "Ё", "Yo" }, { "Ж", "J" }, { "З", "Z" }, { "И", "I" },
                { "Й", "Y" }, { "К", "K" }, { "Л", "L" }, { "М", "M" }, { "Н", "N" },
                { "О", "O" }, { "П", "P" }, { "Р", "R" }, { "С", "S" }, { "Т", "T" },
                { "У", "U" }, { "Ф", "F" }, { "Х", "X" }, { "Ц", "Ts" }, { "Ч", "Ch" },
                { "Ш", "Sh" }, { "Щ", "Sh" }, { "Ъ", "" }, { "Ы", "I" }, { "Ь", "" },
                { "Э", "E" }, { "Ю", "Yu" }, { "Я", "Ya" },
                { "а", "a" }, { "б", "b" }, { "в", "v" }, { "г", "g" }, { "д", "d" },
                { "е", "e" }, { "ё", "yo" }, { "ж", "j" }, { "з", "z" }, { "и", "i" },
                { "й", "y" }, { "к", "k" }, { "л", "l" }, { "м", "m" }, { "н", "n" },
                { "о", "o" }, { "п", "p" }, { "р", "r" }, { "с", "s" }, { "т", "t" },
                { "у", "u" }, { "ф", "f" }, { "х", "x" }, { "ц", "ts" }, { "ч", "ch" },
                { "ш", "sh" }, { "щ", "sh" }, { "ъ", "" }, { "ы", "i" }, { "ь", "" },
                { "э", "e" }, { "ю", "yu" }, { "я", "ya" }
            };

            return string.Join("",_result.Select(c => dict.ContainsKey(c.ToString()) ? dict[c.ToString()] : c.ToString()));
        }
        private string LotinToKrill(string text)
        {
            var _dict = new Dictionary<string, string> {
                {"Yo", "Ё"}, {"Ts", "Ц"}, {"Ch", "Ч"}, {"Sh", "Ш"}, {"Yu", "Ю"}, {"Ya", "Я"},{"O'","Ў"},{"O‘","Ў"},{"G'","Ғ"},{"G‘","Ғ"},
                {"YO", "Ё"}, {"TS", "Ц"}, {"CH", "Ч"}, {"SH", "Ш"}, {"YU", "Ю"}, {"YA", "Я"},{"o'","ў"},{"o‘","ў"},{"g'","ғ"},{"g‘","ғ"},
                {"yo", "ё"}, {"ts", "ц"}, {"ch", "ч"}, {"sh", "ш"}, {"yu", "ю"}, {"ya", "я"},
                {"yO", "ё"}, {"tS", "ц"}, {"cH", "ч"}, {"sH", "ш"}, {"yU", "ю"}, {"yA", "я"}
            };
            var _result = string.Join("", text.Select(c => _dict.ContainsKey(c.ToString()) ? _dict[c.ToString()] : c.ToString()));

            var dict = new Dictionary<string, string>
            {
                { "Q", "Қ" }, { "Ҳ", "H" },
                { "A", "А" }, { "B", "Б" }, { "V", "В" }, { "G", "Г" }, { "D", "Д" },
                { "E", "Е" }, { "Yo", "Ё" }, { "J", "Ж" }, { "Z", "З" }, { "I", "И" },
                { "Y", "Й" }, { "K", "К" }, { "L", "Л" }, { "M", "М" }, { "N", "Н" },
                { "O", "О" }, { "P", "П" }, { "R", "Р" }, { "S", "С" }, { "T", "Т" },
                { "U", "У" }, { "F", "Ф" }, { "X", "Х" }, { "Ts", "Ц" }, { "Ch", "Ч" },
                { "Sh", "Ш" }, { "Yu", "Ю" }, { "Ya", "Я" },
                {"q","қ"},{"ҳ","h"},
                { "a", "а" }, { "b", "б" }, { "v", "в" }, { "g", "г" }, { "d", "д" },
                { "e", "е" }, { "yo", "ё" }, { "j", "ж" }, { "z", "з" }, { "i", "и" },
                { "y", "й" }, { "k", "к" }, { "l", "л" }, { "m", "м" }, { "n", "н" },
                { "o", "о" }, { "p", "п" }, { "r", "р" }, { "s", "с" }, { "t", "т" },
                { "u", "у" }, { "f", "ф" }, { "x", "х" }, { "ts", "ц" }, { "ch", "ч" },
                { "sh", "ш" }, { "yu", "ю" }, { "ya", "я" }
            };
            return new string(_result.Select(c => dict.ContainsKey(c.ToString()) ? dict[c.ToString()][0] : c).ToArray());
        }

    }
}
