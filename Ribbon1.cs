using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace TarjimonWord
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            KrillToLatin();
        }
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            LotinToKrill();
        }

        private void KrillToLatin()
        {
            var _dict = new Dictionary<string, string> {
                { "Ў","O‘" }, { "ў","o‘"},{ "Ғ","G‘" }, { "ғ","g‘" },
                { "Ҳ","H"}, { "ҳ","h"}, { "Қ","Q" }, { "қ","q" }
            };
            _dict.ToList().ForEach(keyValue =>
            {
                SearchReplace(keyValue.Key, keyValue.Value);
            });
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

            dict.ToList().ForEach(keyValue =>
            {
                SearchReplace(keyValue.Key, keyValue.Value);
            });
        }
        private void LotinToKrill()
        {
            var _dict = new Dictionary<string, string> {
                {"Yo", "Ё"}, {"Ts", "Ц"}, {"Ch", "Ч"}, {"Sh", "Ш"}, {"Yu", "Ю"}, {"Ya", "Я"},{"O'","Ў"},{"O‘","Ў"},{"G'","Ғ"},{"G‘","Ғ"},
                {"YO", "Ё"}, {"TS", "Ц"}, {"CH", "Ч"}, {"SH", "Ш"}, {"YU", "Ю"}, {"YA", "Я"},{"o'","ў"},{"o‘","ў"},{"g'","ғ"},{"g‘","ғ"},
                {"yo", "ё"}, {"ts", "ц"}, {"ch", "ч"}, {"sh", "ш"}, {"yu", "ю"}, {"ya", "я"},
                {"yO", "ё"}, {"tS", "ц"}, {"cH", "ч"}, {"sH", "ш"}, {"yU", "ю"}, {"yA", "я"}
            };
            _dict.ToList().ForEach(keyValue =>
            {
                SearchReplace(keyValue.Key, keyValue.Value);
            });
            var dict = new Dictionary<string, string>
            {
                { "Q", "Қ" }, { "H", "Ҳ" },
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
            dict.ToList().ForEach(keyValue =>
            {
                SearchReplace(keyValue.Key, keyValue.Value);
            });
        }
        private void SearchReplace(string findTxt,string replaceTxt)
        {
            object FindTxt = findTxt;
            object ReplaceTxt = replaceTxt;
            Globals.ThisAddIn.Application.ActiveDocument.Content.Find.ClearFormatting();
            Globals.ThisAddIn.Application.ActiveDocument.Content.Find.Execute(ref FindTxt,
                                                     ref Params.matchCase,
                                                     ref Params.matchWholeWord,
                                                     ref Params.matchWildCards,
                                                     ref Params.matchSoundLike,
                                                     ref Params.nmatchAllWordForms,
                                                     ref Params.forward,
                                                     ref Params.wrap,
                                                     ref Params.format,
                                                     ref ReplaceTxt,
                                                     ref Params.replace,
                                                     ref Params.matchKashida,
                                                     ref Params.matchDiacritics,
                                                     ref Params.matchAlefHamza,
                                                     ref Params.matchControl);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("TarjimonWord v1.0\nDeveloped by Developer Temur\nhttps://github.com/ganiyevtemur1/TarjimonWord");
        }
    }
}
