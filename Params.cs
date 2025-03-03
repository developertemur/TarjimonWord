using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TarjimonWord
{
    class Params
    {
        public static object matchCase = true;
        public static object matchWholeWord = false;
        public static object matchWildCards = false;
        public static object matchSoundLike = false;
        public static object nmatchAllWordForms = false;
        public static object forward = true;
        public static object format = false;
        public static object matchKashida = false;
        public static object matchDiacritics = false;
        public static object matchAlefHamza = false;
        public static object matchControl = false;
        public static object read_only = false;
        public static object visible = false;
        public static object replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
        public static object wrap = 1;
    }
}
