using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TileSheet
{
    internal class Constants
    {

        public const string SET_CHARACTERISTICS = "06_Плиточное покрытие"; // поиск набора характеристик

        public static List<string> CHECK_STRINGS = new List<string> { "Единицы измерения", "Количество", "Наименование", "Тип", "Примечание", "id"}; // данные этих строк из НХ нам нужны

        public static Dictionary<int, string> DICT_NAMES_COLUMNS = new Dictionary<int, string>() {
                    { 0, "Тип"},
                    { 1, "Условное обозначение"},
                    { 2, "Наименование"},
                    { 3, "Единицы измерения"},
                    { 4, "Количество"},
                    { 5, "Примечание"}
                };


        public static List<string> NAMES_COLUMNS = new List<string> { "Тип", "Условное обозначение", "Наименование", "Ед.изм.", "Кол-во", "Примечание" };

        public const string TILEKEYWORD = "Плитка";

    }
}
