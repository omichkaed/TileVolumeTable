using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil;
using Autodesk.Civil.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Aec.PropertyData.DatabaseServices;
using static TileSheet.Constants;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using Autodesk.Aec.DatabaseServices;
using ObjectId = Autodesk.AutoCAD.DatabaseServices.ObjectId;
using System.Reflection;
using System.Collections;
using System.Text.RegularExpressions;

namespace TileVolumeTable
{

    public class TileSheet
    {
        [CommandMethod("Tile", "TileVolumeTable", CommandFlags.Modal)]

        public static void TileVolumeTable()
        {

            //   Получаем список словарей со всеми типами используемых плиток в модели
            List<Dictionary<string, string>> tileInfo = GetInfo.TileInfo();
            if (tileInfo == null)
            {
                MessageBox.Show("В пространстве модели сведений не найдено");
                return;
            }
            //   Сортируем всю информацию и в частности прибавляем площади одинаковых штриховок друг к другу
            List<Dictionary<string, string>> sortInfo = GetInfo.SortInfo(tileInfo);

            //   Получаем блоки из чертежа относящиеся к плиточному покрытию
            List<KeyValuePair<string, ObjectId>> blocks = GetInfo.GetBlock();

            Tables.CreateTable(sortInfo, blocks);
        }
    }

    public class GetInfo
    {
        public static List<Dictionary<string, string>> TileInfo()
        {
            List<Dictionary<string, string>> tileInfo = new List<Dictionary<string, string>>();

            // Получение текущего документа
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Начать транзакцию
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Получение модели пространства
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead);
                // Получение ID всех объектов
                ObjectId[] objectIds = modelSpace.Cast<ObjectId>().ToArray();

                foreach (ObjectId objectId in objectIds)
                {
                    Autodesk.AutoCAD.DatabaseServices.DBObject obj = tr.GetObject(objectId, OpenMode.ForRead);
                    if (obj != null)
                    {
                        // Доступ к свойствам набора
                        Autodesk.AutoCAD.DatabaseServices.ObjectIdCollection setIds = PropertyDataServices.GetPropertySets(obj);
                        foreach (ObjectId psId in setIds)
                        {
                            PropertySet pset = tr.GetObject(psId, OpenMode.ForRead) as PropertySet;
                            // Проверяем есть ли нужный нам набор характеристик в этом объекте
                            if (pset.PropertySetDefinitionName == SET_CHARACTERISTICS)
                            {
                                // создаём словарь, в который будут добавляться все нужные нам объекты и их значения с требуемым НХ 
                                Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();

                                PropertySetDataCollection psetDataColl = pset.PropertySetData;

                                PropertySetDefinition propDef = (PropertySetDefinition)tr.GetObject(pset.PropertySetDefinition, OpenMode.ForRead);
                                var psDefCol = propDef.Definitions;


                                foreach (PropertyDefinition psDef in psDefCol)
                                {
                                    if (CHECK_STRINGS.Any(s => s == psDef.Name))
                                    {
                                        var psDefCoId = psDef.Id;
                                        var psDefCoName = psDef.Name;
                                        var dataPSet = pset.GetAt(psDefCoId).ToString();
                                        keyValuePairs.Add(psDefCoName, dataPSet);
                                    }

                                }
                                tileInfo.Add(keyValuePairs);
                            }

                        }
                    }
                }

                tr.Commit();
            }
            if (tileInfo.Count == 0)
                return null;

            return tileInfo;
        }


        public static List<Dictionary<string, string>> SortInfo(List<Dictionary<string, string>> tileInfo)
        {
            List<Dictionary<string, string>> info = new List<Dictionary<string, string>>();
            List<string> tempNameObject = new List<string> { }; // создаем временную строку, в которую добавляться значения, которые впервый раз встретились в принятых данных

            foreach (Dictionary<string, string> elem in tileInfo)
            {
                if (tempNameObject.Any(s => s == elem[CHECK_STRINGS[2]])) // если объект уже есть в временной строке, до далее будем добавлять к этому элементу ещё м2 или шт.
                {
                    foreach (var dict in info)
                    {

                        if (elem[CHECK_STRINGS[2]] == dict[CHECK_STRINGS[2]])
                        {
                            dict[CHECK_STRINGS[1]] = (double.Parse(elem[CHECK_STRINGS[1]]) + double.Parse(dict[CHECK_STRINGS[1]])).ToString();
                        }
                    }
                }
                else // если отсутствует объект, то просто добавляем его во временную строку
                {
                    tempNameObject.Add(elem[CHECK_STRINGS[2]]);
                    info.Add(elem);
                }
            }
            // сортируем наш словарь по типу, который вбил сам проектировщик
            List<Dictionary<string, string>> sortInfo = info.OrderBy(x => Regex.Replace(x["Тип"], @"[^\d\.]", "")).ToList();
            return sortInfo;
        }

        public static List<KeyValuePair<string, ObjectId>> GetBlock()
        {
            List<KeyValuePair<string, ObjectId>> blocks = new List<KeyValuePair<string, ObjectId>>(); // список куда будут включаться все блоки содержащии требуемое ключевое слово.

            // Получение текущего документа
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Начать транзакцию
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {

                // получаем все блоки
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                if (bt != null) 
                {
                    foreach (ObjectId blockId in bt)
                    {
                        // получаем запись кокретно блока
                        BlockTableRecord blockTableRecord = (BlockTableRecord)tr.GetObject(blockId, OpenMode.ForRead);
                        if (blockTableRecord.Name.Contains(TILEKEYWORD)) // если блок содержит ключевое слово, то добавляем в наш беречень блоков
                        {
                            blocks.Add(new KeyValuePair<string, ObjectId>(blockTableRecord.Name, blockId));
                        }

                    }
                }

                
            }
            if (blocks.Count == 0)
                MessageBox.Show("Отсутсвуют исходные блоки плиточного покрытия.\n" +
                    "Для корректного вывода ведомости добавьте блоки плиточного покрытия в чертёж.\n" +
                    "Возможно требуется добавить блоки из менеджера блоков.");
            return blocks;
        }
    }
    public class Tables
    {
        public static void CreateTable(List<Dictionary<string, string>> sortInfo, List<KeyValuePair<string, ObjectId>> blocks)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptPointResult pr = ed.GetPoint("\nУкажите точку вставки таблицы: ");
            if (pr.Status == PromptStatus.OK)
            {

                Autodesk.AutoCAD.DatabaseServices.Table tb = new Autodesk.AutoCAD.DatabaseServices.Table();
                int header = 2;
                int tbSizeColumns = 6;
                int footer = 1;
                int tbSizeRows = header + sortInfo.Count + footer;
                tb.SetSize(tbSizeRows, tbSizeColumns);
                tb.Position = pr.Value;


                //Имя ведомости (первая строка)
                tb.Rows[0].Height = 10;
                tb.Cells[0, 0].TextString = "ВЕДОМОСТЬ ПЛИТОЧНЫХ ПОКРЫТИЙ";
                tb.Cells[0, 0].TextHeight = 3.5;
                CellRange cr = CellRange.Create(tb, 0, 0, 0, tbSizeColumns - 1);
                tb.MergeCells(cr);
                tb.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

                //Последняя строка, частный случай
                tb.Rows[tbSizeRows - 1].Height = 6.5;
                tb.Cells[tbSizeRows - 1, 0].TextString = "Объёмы плитки во входных группах и конструкцию покрытия см. раздел";
                tb.Cells[tbSizeRows - 1, 0].TextHeight = 2.5;
                CellRange ft = CellRange.Create(tb, tbSizeRows - 1, 0, tbSizeRows - 1, tbSizeColumns - 1);
                tb.MergeCells(ft);
                tb.Cells[tbSizeRows - 1, 0].Alignment = CellAlignment.MiddleCenter;

                //Имена столбцов (вторая строка)
                tb.Rows[1].Height = 10;

                tb.Columns[0].Width = 15;
                tb.Cells[1, 0].TextString = NAMES_COLUMNS[0];

                tb.Columns[1].Width = 35;
                tb.Cells[1, 1].TextString = NAMES_COLUMNS[1];

                tb.Columns[2].Width = 110;
                tb.Cells[1, 2].TextString = NAMES_COLUMNS[2];

                tb.Columns[3].Width = 15;
                tb.Cells[1, 3].TextString = NAMES_COLUMNS[3];

                tb.Columns[4].Width = 15;
                tb.Cells[1, 4].TextString = NAMES_COLUMNS[4];

                tb.Columns[5].Width = 35;
                tb.Cells[1, 5].TextString = NAMES_COLUMNS[5];

                for (int col = 0; col < tbSizeColumns; col++)
                {
                    tb.Cells[1, col].TextHeight = 2.5;
                    tb.Cells[1, col].Alignment = CellAlignment.MiddleCenter;

                }

                // оформление таблицы

                for (int row = header; row < tbSizeRows - footer; row++)
                {

                    tb.Rows[row].Height = 6.5;
                    for (int col = 0; col < tbSizeColumns; col++)
                    {
                        tb.Cells[row, col].TextHeight = 2.5;
                        if (sortInfo[row - header].ContainsKey(DICT_NAMES_COLUMNS[col]))
                            tb.Cells[row, col].TextString = sortInfo[row - header][DICT_NAMES_COLUMNS[col]];
                        if (col != 2)
                        {
                            tb.Cells[row, col].Alignment = CellAlignment.MiddleCenter;
                            if (col == 1)
                            {
                                string temp = sortInfo[row - header][CHECK_STRINGS[5]];
                                int cnt = blocks.Count;
                                foreach (KeyValuePair<string, ObjectId> block in blocks)
                                {
                                    if (block.Key.Contains(temp))
                                    {
                                        tb.Cells[row, col].BlockTableRecordId = block.Value;
                                        tb.Cells[row, col].Contents[0].IsAutoScale = true;
                                        //tb.Cells[row, col].Contents[0].Scale = 4;
                                        break;
                                    }
                                    else
                                        cnt--;
                                }
                                if (cnt == 0)
                                    MessageBox.Show($"Блок {temp} отсутствует в чертеже.\n" +
                    $"Для корректного вывода ведомости добавьте блок {temp} плиточного покрытия в чертёж.\n" +
                    "Возможно требуется добавить блоки из менеджера блоков.");

                            }
                        }
                        else
                            tb.Cells[row, col].Alignment = CellAlignment.MiddleLeft;

                    }
                }

                tb.GenerateLayout();
                Transaction tr = doc.TransactionManager.StartTransaction();

                using (tr)
                {
                    BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    btr.AppendEntity(tb);
                    tr.AddNewlyCreatedDBObject(tb, true);
                    tr.Commit();
                }
            }
            else
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nВыполнение прервано");
                return;
            }
        }
    }
}

