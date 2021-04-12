using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace parsing
{
    class Program
    {
        public static List<Entity> Entities = new List<Entity>();
        public static List<Relation> Relations = new List<Relation>();
        public static int[] used_rules = new int[6] { 0, 0, 0, 0, 0, 0 };
        public static string[][] headers = new string[2][]
            {
                new string[6] {"Имя атрибута", "Назначение атрибута", "Тип", "Формат",  "Допустимые значения", "Примечание" },
                new string[3] {"Сущность – связь – сущность","Применяемое правило", "Получаемое реляционное отношение"}
            };

        public class Entity
        {
            public string Name;
            public string Additional;
            public List<Attribute> Attributes = new List<Attribute>();
            public Entity(string name)
            {
                Name = name;
            }
        }
        public class Attribute
        {
            public string Name;
            public string Type;
            public string Format;
            public string Info;
            public bool Primary;
            public int Foreign;
            public bool Nullable;
            public int Alternative;
            public Attribute(string name, string type, string format, string info, bool key, bool nullable)
            {
                Name = name;
                Type = type;
                Format = format;
                Info = info;
                Foreign = 0;
                Alternative = 0;
                Primary = key;
                Nullable = nullable;
            }
            public Attribute(Attribute a)
            {
                Name = a.Name;
                Type = a.Type;
                Format = a.Format;
                Info = a.Info;
                Foreign = 0;
                Alternative = 0;
                Primary = false;
                Nullable = a.Nullable;
            }
        }
        public class Relation
        {
            public string Name;
            public Relation_Type[] Relation_type;
            public int rule;
            public Entity First_Entity;
            public Entity Second_Entity;
            public string Additional;
            public Entity[] Rel_Entities;
            public List<Attribute> Attributes;
            private void SetAll(Entity first, Entity second, Relation_Type type, 
                Relation_Type mandatory, string name)
            {
                First_Entity = first;
                Second_Entity = second;
                Relation_type = new Relation_Type[2] { type, mandatory };
                Attributes = new List<Attribute>();
                rule = SetRule();
                Name = name;
            }
            private int SetRule()
            {
                if (Relation_type[0] == Relation_Type.M_M)
                    return 6;
                if (Relation_type[0] == Relation_Type.O_O && Relation_type[1] == Relation_Type.M_M)
                    return 1;
                if (Relation_type[0] == Relation_Type.O_O && Relation_type[1] == Relation_Type.O_O)
                    return 3;
                if (Relation_type[0] == Relation_Type.O_O)
                    return 2;
                if ((Relation_type[0] == Relation_Type.O_M || Relation_type[0] == Relation_Type.M_O) &&
                    (Relation_type[1] == Relation_type[0] || Relation_type[1] == Relation_Type.M_M))
                    return 4;
                return 5;
            }
            public Relation(Entity first, Entity second, Relation_Type type, 
                Relation_Type mandatory, string name)
            {
                SetAll(first, second, type, mandatory, name);
            }
            public Relation(Entity first, Entity second, Relation_Type type, 
                Relation_Type mandatory, string name, string info)
            {
                SetAll(first, second, type, mandatory, name);
                Additional = info;
            }
            public void Rule()
            {
                used_rules[rule - 1]++;
                switch (rule)
                {
                    case 1:
                        {
                            Rel_Entities = new Entity[1] { First_Entity };
                            int alternative = Second_Entity.Attributes.Max(a => a.Alternative) + 1;
                            foreach (Attribute attribute in Second_Entity.Attributes)
                            {
                                Rel_Entities[0].Attributes.Remove(attribute);
                                if (attribute.Primary)
                                    Rel_Entities[0].Attributes.Add(new Attribute(attribute) { Alternative = alternative });
                                else
                                    Rel_Entities[0].Attributes.Add(new Attribute(attribute));
                            }
                            Entities.Remove(Second_Entity);
                            break;
                        }
                    case 2:
                        {
                            if (Relation_type[1] == Relation_Type.M_O)
                                Rel_Entities = new Entity[2] { First_Entity, Second_Entity };
                            else
                                Rel_Entities = new Entity[2] { Second_Entity, First_Entity };
                            Transfer_Attributes(Rel_Entities[0], Rel_Entities[1], false,
                                Rel_Entities[1].Attributes.Max(a => a.Alternative) + 1, 0);
                            break;
                        }
                    case 3:
                        {
                            Linking_entity(false);
                            Transfer_Attributes(Rel_Entities[0], Rel_Entities[2], true, 0, 1);
                            Transfer_Attributes(Rel_Entities[1], Rel_Entities[2], false,
                                Rel_Entities[2].Attributes.Max(a => a.Alternative) + 1, 2);
                            Entities.Add(Rel_Entities[2]);
                            break;
                        }
                    case 4:
                        {
                            if (Relation_type[0] == Relation_Type.O_M)
                                Rel_Entities = new Entity[2] { First_Entity, Second_Entity };
                            else
                                Rel_Entities = new Entity[2] { Second_Entity, First_Entity };
                            int foreign = Rel_Entities[1].Attributes.Max(a => a.Foreign) + 1;
                            Transfer_Attributes(Rel_Entities[0], Rel_Entities[1], false, 0, foreign);
                            break;
                        }
                    case 5:
                        {
                            bool reverse = Relation_type[0] == Relation_Type.M_O;
                            Linking_entity(reverse);
                            Transfer_Attributes(Rel_Entities[0], Rel_Entities[2], false, 0, 1);
                            Transfer_Attributes(Rel_Entities[1], Rel_Entities[2], true, 0, 2);
                            Entities.Add(Rel_Entities[2]);
                            break;
                        }
                    case 6:
                        {
                            Linking_entity(false);
                            Transfer_Attributes(Rel_Entities[0], Rel_Entities[2], true, 0, 1);
                            Transfer_Attributes(Rel_Entities[1], Rel_Entities[2], true, 0, 2);
                            Entities.Add(Rel_Entities[2]);
                            break;
                        }
                    default: Console.WriteLine("HMMMMMMMMMMMMMMM"); break;
                }
            }
            private void Linking_entity(bool reverse)
            {
                Rel_Entities = new Entity[3];
                Entity Additional = new Entity("");
                Additional.Attributes.AddRange(Attributes);
                if (reverse)
                    (Rel_Entities[0], Rel_Entities[1], Rel_Entities[2]) = (Second_Entity, First_Entity, Additional);
                else
                    (Rel_Entities[0], Rel_Entities[1], Rel_Entities[2]) = (First_Entity, Second_Entity, Additional);
                Additional.Name = $"{Rel_Entities[0].Name} -- {Name} -- {Rel_Entities[1].Name}";
            }
            private void Transfer_Attributes(Entity from, Entity to, bool primary, 
                int alternative, int foreign)
            {
                List<Attribute> attributes = from.Attributes.FindAll(a => a.Primary);
                foreach (Attribute attribute in attributes)
                {
                    to.Attributes.Remove(
                    //    to.Attributes.Find(a => a.Name == attribute.Name));
                          attribute);
                    //if (to.Attributes.Find(a => a.Name == attribute.Name) == null)
                    to.Attributes.Add(new Attribute(attribute)
                    {
                        Primary = primary,
                        Alternative = alternative,
                        Foreign = foreign
                    });
                }
            }
            public string First_Column()
            {
                string O_M, man_opt;
                if ((int)Relation_type[0] % 2 == 1)
                    O_M = "1 - ";
                else
                    O_M = "М - ";
                if ((int)Relation_type[0] > 20)
                    O_M += "М";
                else
                    O_M += "1";
                if ((int)Relation_type[1] % 2 == 1)
                    man_opt = "Необ. - ";
                else
                    man_opt = "Об. - ";
                if ((int)Relation_type[1] > 20)
                    man_opt += "об.";
                else
                    man_opt += "необ.";
                return (First_Entity.Name + " - " + Name + " - " + Second_Entity.Name + 
                    Environment.NewLine + O_M + Environment.NewLine + man_opt);
            }
            public string Third_Column()
            {
                string result = "";
                foreach (Entity e in Rel_Entities)
                {
                    result += $"{e.Name}(";
                    e.Attributes.ForEach(a => result += a.Name + ", ");
                    result = $"{result.Substring(0, result.Length - 2)}){Environment.NewLine}Первичный ключ: ";
                    e.Attributes.ForEach(a => { if (a.Primary) result += $"{a.Name}, "; });
                    result = result.Substring(0, result.Length - 2) + Environment.NewLine;
                    {
                        List<Attribute> alt = e.Attributes.FindAll(a => a.Alternative > 0);
                        if (alt.Any())
                            for (int a = 1; a < alt.Max(m => m.Alternative); a++)
                            {
                                result += "Альтернативный ключ №" + a + ": ";
                                alt.FindAll(m => m.Alternative == a).
                                    ForEach(act => result += act.Name + ", ");
                                result = result.Substring(0, result.Length - 2) + Environment.NewLine;
                            }
                    }
                    {
                        List<Attribute> foreign = e.Attributes.FindAll(a => a.Foreign > 0);
                        if (foreign.Any())
                            for (int a = 1; a <= foreign.Max(m => m.Foreign); a++)
                            {
                                result += "Внешний ключ №" + a + ": ";
                                foreign.FindAll(m => m.Foreign == a).
                                    ForEach(act => result += act.Name + ", ");
                                result = result.Substring(0, result.Length - 2) + Environment.NewLine;
                            }
                    }
                }
                result = result.Substring(0, result.LastIndexOf(Environment.NewLine));
                return result;
            }


        }
        public enum Relation_Type
        {
            O_O = 3,    // One to One, or Optional to Optional
            M_O = 12,   // Many to One, ot Mandatory to Optional
            O_M = 21,   // One to Many, or Optional to Mandatory
            M_M = 30    // Many to Many, or Mandatory to Mandatory
        }
        static void Main()
        {
            string Relations_dir = null;
            string Entities_dir = null;

            Console.WriteLine("Enter names of 2 txt files (UTF-8)");
            do
            {
                if (string.IsNullOrWhiteSpace(Entities_dir))
                {
                    Console.WriteLine("Enter name of files with entities (full path, if in different directories with .exe file):");
                    Entities_dir = Console.ReadLine();
                    if (!File.Exists(Entities_dir))
                        Entities_dir = null;
                }
                if (string.IsNullOrWhiteSpace(Relations_dir))
                {
                    Console.WriteLine("Enter name of file with relations:");
                    Relations_dir = Console.ReadLine();
                    if (!File.Exists(Relations_dir))
                        Relations_dir = null;
                }
            } while (Entities_dir == null || string.IsNullOrWhiteSpace(Relations_dir));

            Extract_Entities(Entities_dir);
            Extract_Relations(Relations_dir);

            int go = 5;
            while (go > 0)
            {
                Console.WriteLine("What to do?\n" +
                    "1: Entities to Excel;\n" +
                    "2: Relations to Excel;\n" +
                    "3: Entities to Word;\n" +
                    "4: Relations to Word");
                int.TryParse(Console.ReadLine(), out go);
                Console.WriteLine("Enter new file's name");
                string filename = Console.ReadLine();
                switch (go)
                {
                    case 1: Entities_to_Excel(filename); break;
                    case 2: Relations_to_Excel(filename); break;
                    case 3: Entities_to_Word(filename); break;
                    case 4: Relations_to_Word(filename); break;
                    default: Console.WriteLine("Wrong command input, 1-4 please"); continue;
                }
            }
        }
        private static void Extract_Relations(string directory)
        {
            List<string> RList = new StreamReader(directory).ReadToEnd().
                Split(new char[2] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            bool LastAddAttribute = false;
            foreach (string s in RList)
            {
                if (s.Contains("REGULAR RELATION"))
                {
                    string[] names = new string[2];
                    int type = 0, mandatory = 0;
                    string name = s.Substring(0, s.IndexOf(':') - 1);
                    string[] entities_in_strings = s.Substring(s.IndexOf("RELATIONSHIP") + 13).
                        Split(new string[1] { " to " }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < 2; i++)
                    {
                        int one_index = entities_in_strings[i].IndexOf("ONE");
                        int many_index = entities_in_strings[i].IndexOf("MANY");

                        int optional_index = entities_in_strings[i].IndexOf("OPTIONAL");
                        if (one_index > -1)
                        {
                            type += 1 * (i + 1);
                            names[i] = entities_in_strings[i].Substring(0, one_index - 1);
                        }
                        else
                        {
                            type += 10 * (i + 1);
                            names[i] = entities_in_strings[i].Substring(0, many_index - 1);
                        }
                        if (optional_index > -1)
                            mandatory += 1 * (i + 1);
                        else
                            mandatory += 10 * (i + 1);
                    }

                    Relations.Add(new Relation(Entities.Find(e => e.Name.Equals(names[0])),
                        Entities.Find(e => e.Name.Equals(names[1])), (Relation_Type)type, (Relation_Type)mandatory, name));
                    LastAddAttribute = false;
                }
                else
                {
                    if (s.Contains("ATTRIBUTE"))
                    {
                        AddAttribute(Relations.Last().Attributes, s);
                        LastAddAttribute = true;
                    }
                    else
                    {
                        if (LastAddAttribute)
                            Relations.Last().Attributes.Last().Info = s;
                        else
                            Relations.Last().Additional = s;
                    }
                }
            }
            Relations.ForEach(r => r.Rule());
        }
        private static void Relations_to_Excel(string filename)
        {
            int count = Relations.Count + used_rules[2] + used_rules[4] + used_rules[5];
            Excel.Application app = new Excel.Application { SheetsInNewWorkbook = count };
            Excel.Workbooks workbooks = app.Workbooks;
            Excel.Workbook workbook = workbooks.Add();

            int k = 1;
            foreach (Relation r in Relations)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets[k];
                sheet.Cells[1, 1] = r.Name;
                sheet.Cells[1, 2] = r.Additional;
                for (int i = 0; i < headers[1].Length; i++)
                    sheet.Cells[2, i + 1] = headers[1][i];

                sheet.Cells[3, 1] = r.First_Column();
                sheet.Cells[3, 2] = $"Правило {r.rule}";
                sheet.Cells[3, 3] = r.Third_Column();
                Marshal.ReleaseComObject(sheet);
                k++;
            }
            workbook.SaveAs(filename + ".xlsx");
            app.Visible = true;
        }
        private static void Extract_Entities(string directory)
        {
            string Entities_strings = new StreamReader(directory).ReadToEnd();

            List<string> Names = (from s in Entities_strings.Split('\n')
                                  where s.Contains("REGULAR ENTITY")
                                  select s.Substring(0, s.Length - 17).Trim('\n', '\r', ' ')).ToList();

            int index, next_index = 0;
            index = Entities_strings.IndexOf("REGULAR ENTITY", next_index) + 16;
            next_index = Entities_strings.IndexOf("REGULAR ENTITY", index);

            List<string> attributes = new List<string>();
            bool run = true;
            do
            {
                if (next_index < index)
                {
                    next_index = Entities_strings.Length - 1;
                    run = false;
                }
                string s = Entities_strings.Substring(index, next_index - index - 2);
                attributes.Add(s.Substring(0, s.LastIndexOf('\r')));

                index = Entities_strings.IndexOf("REGULAR ENTITY", next_index) + 16;
                next_index = Entities_strings.IndexOf("REGULAR ENTITY", index);
            } while (run);

            for (int i = 0; i < Names.Count; i++)
            {
                Entities.Add(new Entity(Names[i].Trim('\r')));
                List<string> attributes_i = attributes[i].Split(new char[2] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).ToList();

                if (!attributes_i.First().Contains("ATTRIBUTE"))
                {
                    Entities.Last().Additional = attributes_i.First();
                    attributes_i.RemoveAt(0);
                }
                foreach (string attribute in attributes_i)
                    AddAttribute(Entities.Last().Attributes, attribute);
            }
        }
        public static void AddAttribute(List<Attribute> Attributes, string attribute)
        {
            string name, type = "", format = "", info = "";
            bool nullable = false, primary_key = false;
            if (!attribute.Contains("ATTRIBUTE"))
            {
                Attributes.Last().Info = attribute;
                return;
            }
            if (attribute.Contains("[PK]"))
                primary_key = true;
            if (attribute.Contains("[ALLOW NULL]"))
                nullable = true;
            {
                if (attribute.Contains("Type:CHAR") || attribute.Contains("Type:VARCHAR"))
                {
                    type = "Символьный";
                    if (attribute.Contains("Length="))
                    {
                        int start = attribute.IndexOf("Length=") + 7;
                        string length = attribute.Substring(start, attribute.Length - start);
                        if (!string.IsNullOrEmpty(length))
                            format = $"CHAR (Длина - {length})";
                    }
                }
                else
                if (attribute.Contains("Type:INTEGER"))
                    type = "Числовой";
                else
                if (attribute.Contains("Type:DATE"))
                    type = "Дата";
                else
                if (attribute.Contains("Type:NO DATATYPE"))
                    type = "NO DATATYPE";
            }               // type, format
            string search_item = " : ATTRIBUTE";
            if (primary_key)
                search_item = " : [PK]";
            else
                if (nullable)
                search_item = " : [ALLOW NULL]";

            name = attribute.Substring(0, attribute.IndexOf(search_item));

            Attributes.Add(new Attribute(name, type, format, info, primary_key, nullable));

        }
        private static void Entities_to_Excel(string filename)
        {
            Excel.Application app = new Excel.Application { SheetsInNewWorkbook = Entities.Count };
            Excel.Workbooks workbooks = app.Workbooks;
            Excel.Workbook workbook = workbooks.Add();
            int k = 1;
            foreach (Entity entity in Entities)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(k);
                sheet.Cells[1, 1] = entity.Name;
                sheet.Cells[1, 2] = entity.Additional;
                for (int i = 0; i < headers[0].Length; i++)
                    sheet.Cells[2, i + 1] = headers[0][i];
                for (int j = 0; j < entity.Attributes.Count; j++)
                {
                    sheet.Cells[j + 3, 1] = entity.Attributes[j].Name;
                    sheet.Cells[j + 3, 3] = entity.Attributes[j].Type;
                    sheet.Cells[j + 3, 4] = entity.Attributes[j].Format;
                    sheet.Cells[j + 3, 6] = GetInfo(entity.Attributes[j]);
                }
                sheet = null;
                k++;
            }
            workbook.SaveAs(filename + ".xlsx");
            app.Visible = true;
        }
        private static string GetInfo(Attribute a)
        {
            string info = a.Info;
            if (!string.IsNullOrWhiteSpace(info))
                info += "; ";
            if (a.Primary)
                info += "первичный ключ; ";
            if (a.Foreign > 0)
                info += "внешний ключ №" + a.Foreign + "; ";
            if (a.Alternative > 0)
                info += "альтернативный ключ №" + a.Alternative + "; ";
            if (info.StartsWith(Environment.NewLine))
                info = info.Trim(' ');
            return info;
        }
        private static void  Entities_to_Word(string filename)
        {
            Word.Application app = new Word.Application();
            Word.Document word = app.Documents.Add();
            Word.Table table = word.Tables.Add(word.Range(), 
                Entities.Sum(e => e.Attributes.Count) + 3 * Entities.Count, 6);
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            int e_count = 0;
            int row_count = 1;
            foreach (Entity e in Entities)
            {
                table.Cell(row_count, 1).Range.Text = e.Name;
                table.Cell(row_count, 1).Range.Font.Bold = 1;
                if (!string.IsNullOrWhiteSpace(e.Additional))
                {
                    table.Cell(row_count, 2).Range.Text = e.Additional;
                    table.Cell(row_count, 2).Merge(table.Cell(row_count, 6));
                }
                else
                    table.Cell(row_count, 1).Merge(table.Cell(row_count, 6));
                row_count++;
                for (int i = 0; i < headers[0].Length; i++)
                    table.Cell(row_count, i + 1).Range.Text = headers[0][i];
                row_count++;
                foreach (Attribute a in e.Attributes)
                {
                    table.Cell(row_count, 1).Range.Text = a.Name;
                    table.Cell(row_count, 3).Range.Text = a.Type;
                    table.Cell(row_count, 4).Range.Text = a.Format;
                    table.Cell(row_count, 6).Range.Text = GetInfo(a);
                    row_count++;
                }
                table.Cell(row_count, 1).Merge(table.Cell(row_count, 6));
                row_count++;
                e_count++;
            }
            table.Rows.Last.Delete();
            word.SaveAs(filename + ".docx");
            app.Visible = true;
        }
        private static void Relations_to_Word(string filename) 
        {
            Word.Application app = new Word.Application();
            Word.Document word = app.Documents.Add();
            Word.Table table = word.Tables.Add(word.Range(), Relations.Count * 4, 3);
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            int row_count = 1;
            foreach(Relation r in Relations)
            {
                table.Cell(row_count, 1).Range.Text = r.Name;
                table.Cell(row_count, 1).Range.Font.Bold = 1;
                if (!string.IsNullOrWhiteSpace(r.Additional))
                {
                    table.Cell(row_count, 2).Merge(table.Cell(row_count, 3));
                    table.Cell(row_count, 2).Range.Text = r.Additional;
                }
                else
                    table.Cell(row_count, 1).Merge(table.Cell(row_count, 3));
                row_count++;
                for (int i = 0; i < headers[1].Length; i++)
                    table.Cell(row_count, i + 1).Range.Text = headers[1][i];
                row_count++;
                table.Cell(row_count, 1).Range.Text = r.First_Column();
                table.Cell(row_count, 2).Range.Text = $"Правило {r.rule}";
                table.Cell(row_count, 3).Range.Text = r.Third_Column();
                row_count++;
                table.Cell(row_count, 1).Merge(table.Cell(row_count, 3));
                row_count++;
            }
            table.Rows.Last.Delete();
            word.SaveAs(filename + ".docx");
            app.Visible = true;
        }
    }
}