using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Emionov_root;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;




namespace Eminem
{

    class Program
    {
       static  Word.Application application;
        static Word.Document document;
        static Object missingObj = System.Reflection.Missing.Value;
        static Object trueObj = true;
        static Object falseObj = false;
        static string FilePatch = Directory.GetCurrentDirectory()+"\\otchet.docx";
        static string SaveFilePatch = Directory.GetCurrentDirectory()+"\\@otchet";
        static string CabelDlinaFloor;
        static string BetweenZdaniiCabel;
        static int tables_count = 0;


        static void Main(string[] args)
        {
            /*try
            {*/

                int itog_dlina_ethernet = 0;
                int itog_dlina_optovolokno = 0;
                int itog_dlina_koaks = 0;
                int cab_can_dlina_itog = 0;
                


                Open_doc();


                Console.WriteLine("Введите номер вашего варианта, первая цифра это последняя цифра номера группы, вторая и третья цифра это ваш вариант. Например, 901:");
                int nom_var = Convert.ToInt32(Console.ReadLine());
                Replase("VariantKurs", Convert.ToString(nom_var));

                Console.WriteLine("Введите количество используемых Информационных систем:");
                int Tech_number = Convert.ToInt32(Console.ReadLine());
                Technology[] T = new Technology[Tech_number];
                string[] tehnol_tab_header= new string[3];
                tehnol_tab_header[0] = "Решаемая задача";
                tehnol_tab_header[1] = "Информационная система";
                tehnol_tab_header[2] = "Характер решаемой задачи";
                string[,] tehno_tab = new string[Tech_number, 3];
                for (int k = 0; k < Tech_number; k++)
                {

                    Console.Write("Введите название Информационной системы № ");
                    Console.Write(k + 1);
                    Console.WriteLine(":");
                    T[k].name = Console.ReadLine();
                    tehno_tab[k, 1] = T[k].name;
                    Console.Write("Введите количество трафика Информационной системы в мб/с от пользователя (только целое число) ");
                    Console.Write(T[k].name);
                    Console.WriteLine(":");
                    T[k].load = Convert.ToInt32(Console.ReadLine());

                }
                ReplaseTable("@@system_table", tehno_tab, tehnol_tab_header,3);
                void pokaz_techno()
                {
                    Console.WriteLine("Напоминание, под каким номером какая Информационная система:");
                    for (int k = 0; k < Tech_number; k++)
                    {

                        Console.Write("Информационная система № ");
                        Console.Write(k + 1);
                        Console.Write(":");
                        Console.WriteLine(T[k].name);

                    }

                }


                Builds A = new Builds();
                Console.WriteLine("Введите количество зданий:");
                int Count = Convert.ToInt32(Console.ReadLine());
                Replase("Kolzdanii", Convert.ToString(Count));
                /*
               Console.WriteLine("Введите расстояния между зданиями, через запятую. Например: 400,600:");
               string rofel = Console.ReadLine();
               Replase("BetweenBuilds", rofel);
               Console.WriteLine("Введите количество этажей в зданиях, через запятую. Например: 4,3:");
               string rofel1 = Console.ReadLine();
               Replase("FloorBuilds", rofel1);
               Console.WriteLine("Введите площадь этажей в зданиях, через запятую. Например: 400,300:");
               string rofel2 = Console.ReadLine();
               Replase("FloorSquare", rofel2);
               Console.WriteLine("Введите высоту этажей в зданиях, через запятую. Например: 4,3:");
               string rofel3 = Console.ReadLine();
               Replase("FloorHeight", rofel3);

               Console.WriteLine("Введите количество рабочих по этажам, через запятую. Например: 400 на первом,300 на втором,200 на остальных:");
               string rofel5 = Console.ReadLine();
               Replase("WorkersFloors", rofel5);
               Console.WriteLine("Введите количество мобильных станций в помещении. Например, 40:");
               string rofel6 = Console.ReadLine();
               Replase("MobStations", rofel6);
               */

                /*

               */
                int[,] matrix = new int[Count, Count];
                Console.WriteLine("Сколько существует пар зданий связанных? ");
                int pari = Convert.ToInt32(Console.ReadLine());
                string betw_buil = "";
                for (int z = 0; z < pari; z++)
                {
                Console.WriteLine("Введите номер первого из двух связанных зданий");
                int first_build = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Введите номер второго из двух связанных зданий");
                int second_build = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Введите расстояние между этими 2 зданиями");
                int length = Convert.ToInt32(Console.ReadLine());
                    matrix[first_build - 1, second_build - 1] = length;
                    matrix[second_build - 1, first_build - 1] = length;
                    betw_buil += first_build + " - " + second_build + " = " + length + ", ";
                }
                Replase("BetweenBuilds", betw_buil);

                string[] depart_tech_haeders = new string[5];
                depart_tech_haeders[0] = "Название отдела";
                depart_tech_haeders[1] = "Способ взаимодействия";
                depart_tech_haeders[2] = "Информационная система";
                depart_tech_haeders[3] = "Организационные единицы";
                depart_tech_haeders[4] = "Характер решаемой задачи";
                int department_count = 0;
                string[,] depart_tech = new string[department_count + 1, 5];
                Connect M = new Connect();
                M.connect_length = matrix;
                int i,all_workers=0;
                A.build = new Build[Count];
                string floor_num = "", squre_num = "", hight_num = "", workers_num = "", mob_st_num = "", descr_otd = "", use_eq = "", uninterruptedpower = "";
                for (i = 0; i < Count; i++)
                {
                    descr_otd += "^p В здании №" + i + "находятся: ";
                    A.build[i] = new Build();
                    Console.Write("Введите количество этажей здания № ");
                    Console.Write(i + 1);
                    Console.WriteLine(":");
                    A.build[i].floor_num = Convert.ToInt32(Console.ReadLine());
                    floor_num += A.build[i].floor_num + ",";
                    Console.Write("Введите площадь этажа здания № ");
                    Console.Write(i + 1);
                    Console.WriteLine(":");
                    A.build[i].square = Convert.ToInt32(Console.ReadLine());
                    squre_num += A.build[i].square + ",";
                    Console.Write("Введите высоту этажа здания № ");
                    Console.Write(i + 1);
                    Console.WriteLine(":");
                    A.build[i].height = Convert.ToInt32(Console.ReadLine());
                    hight_num += A.build[i].height;
                    Console.Write("Введите количество мобильных станций ");
                    Console.Write(" Здания № ");
                    Console.Write(i + 1);
                    Console.WriteLine(":");
                    int mob_st = Convert.ToInt32(Console.ReadLine());
                    mob_st_num += mob_st + ",";
                    workers_num += i + " ";

                    A.build[i].floor = new Floor[A.build[i].floor_num];
                    for (int i1 = 0; i1 < A.build[i].floor_num; i1++)
                    {
                        descr_otd += i1 + "этаж - ";
                        int workers = 0;
                        A.build[i].floor[i1] = new Floor();
                        A.build[i].floor[i1].Mob_st = mob_st;
                        Console.Write("Введите количество отделов этажа № ");
                        Console.Write(i1 + 1);
                        Console.Write(" Здания № ");
                        Console.Write(i + 1);
                        Console.WriteLine(":");
                        A.build[i].floor[i1].dep_num = Convert.ToInt32(Console.ReadLine());
                        A.build[i].floor[i1].dep = new Department[A.build[i].floor[i1].dep_num];
                        for (int i2 = 0; i2 < A.build[i].floor[i1].dep_num; i2++)
                        {
                            pokaz_techno();
                            Console.Write("Введите название отдела № ");
                            Console.Write(i2 + 1);
                            Console.Write(" этажа № ");
                            Console.Write(i1 + 1);
                            Console.Write(" здания № ");
                            Console.Write(i + 1);
                            Console.WriteLine(":");
                            A.build[i].floor[i1].dep[i2].name = Console.ReadLine();
                            depart_tech[department_count, 0] = A.build[i].floor[i1].dep[i2].name;
                            Console.Write("Введите количество работников отдела № ");
                            Console.Write(i2 + 1);
                            Console.Write(" этажа № ");
                            Console.Write(i1 + 1);
                            Console.Write(" здания № ");
                            Console.Write(i + 1);
                            Console.WriteLine(":");
                            A.build[i].floor[i1].dep[i2].workers = Convert.ToInt32(Console.ReadLine());
                            workers += A.build[i].floor[i1].dep[i2].workers;
                            all_workers += A.build[i].floor[i1].dep[i2].workers;
                            Console.Write("Введите количество используемых Информационных систем отдела № ");
                            Console.Write(i2 + 1);
                            Console.Write(" этажа № ");
                            Console.Write(i1 + 1);
                            Console.Write(" здания № ");
                            Console.Write(i + 1);
                            Console.WriteLine(":");
                            A.build[i].floor[i1].dep[i2].techno_count = Convert.ToInt32(Console.ReadLine());
                            A.build[i].floor[i1].dep[i2].tech = new Technology[A.build[i].floor[i1].dep[i2].techno_count];
                            for (int i3 = 0; i3 < A.build[i].floor[i1].dep[i2].techno_count; i3++)
                            {

                                Console.WriteLine("Введите номер Информационной системы, которая используется в отделе:");
                                A.build[i].floor[i1].dep[i2].tech[i3] = T[Convert.ToInt32(Console.ReadLine()) - 1];
                                depart_tech[department_count, 2] += A.build[i].floor[i1].dep[i2].tech[i3].name + "^P"; 
                                A.build[i].floor[i1].dep[i2].tech[i3].user = true;
                                depart_tech[department_count, 1] += "Пользователь";
                                Console.WriteLine("Является ли отдел корневым отделом, к которому идут все запросы? Да - '1' Нет - '2': ");
                                string qw = Console.ReadLine();
                                if (qw == "1")
                                {
                                    A.build[i].floor[i1].dep[i2].tech[i3].root = true;
                                    depart_tech[department_count, 1] += ",Сервер";
                                }
                                else
                                if (qw == "2")
                                    A.build[i].floor[i1].dep[i2].tech[i3].root = false;
                                depart_tech[department_count, 1] += "^i";
                                Console.WriteLine("Использует ли программа ресурсы сети интернет, если да - все запросы считаются внесетевыми и не рассчитываются между зданиями. Да - '1' Нет - '2': ");
                                qw = Console.ReadLine();
                                if (qw == "1")
                                    A.build[i].floor[i1].dep[i2].tech[i3].rem_serv = true;
                                else
                                if (qw == "2")
                                    A.build[i].floor[i1].dep[i2].tech[i3].rem_serv = false;
                            }
                            descr_otd += A.build[i].floor[i1].dep[i2].name + " (" + A.build[i].floor[i1].dep[i2].workers + " рабочих станций), ";
                            department_count++;
                            string[,] buf_mas = new string[department_count + 1, 5];
                            Array.Copy(depart_tech, buf_mas, depart_tech.Length);
                            depart_tech = buf_mas;

                        }
                        workers_num += workers + ",";
                    }
                    workers_num += "^l";
                }
                Console.WriteLine("Подождите, идет заполнение документа");
                ReplaseTable("@@depart_tech", depart_tech, depart_tech_haeders,6);
                //вывод в таблицу
                Replase("FloorsBuilds", floor_num);
                Replase("FloorSquare", squre_num);
                Replase("FloorHeight", hight_num);
                Replase("WorkersFloors", workers_num);
                Replase("MobStations", mob_st_num);
                Console.WriteLine("Введите время реакции системы в мс. Например, 400:");
                string rofel7 = Console.ReadLine();
                Replase("ReactionTime", rofel7);
                Replase("WorkersBuilds", all_workers + " ");

                Replase("@@Otdel", descr_otd);// заполняем описания отделов
                int load_num = 0;

                Method.build_load_calculation(ref A);
                for (int v = 0; v < A.load.Length; v++)
                {
                    load_num += A.load[v];
                    Console.WriteLine("Предполагаемая нагрузка внутри здания № " + (v+1) + " =" + A.load[v]);
                    if (A.load[v] < 1000)
                    {
                        A.build[v].cabel_type = "ethernet";
                        Console.WriteLine("В здании № " + (v+1) + " предполагается использование кабеля ethernet");

                        if (A.build[v].height * A.build[v].floor_num > 100)
                            Console.WriteLine("Требуется применение дополнительных устройств усиления для вертикальных кабелей");
                        Console.WriteLine("Вы согласны? 1 - Да, 2 - Нет ");

                        if (Console.ReadLine() == "2")
                        {
                            A.build[v].cabel_type = "оптоволокно";
                            for (int i5 = 0; i5 < A.build[v].floor_num; i5++)
                                itog_dlina_optovolokno += A.build[v].height * i5;

                        }
                        else
                            for (int i5 = 0; i5 < A.build[v].floor_num; i5++)
                                itog_dlina_ethernet += A.build[v].height * i5;
                    }
                    else
                    {
                        Console.WriteLine("В здании № " + (v + 1) + " предполагается использование оптоволоконного кабеля, поскольку нагрузка на сеть больше 1000 Мб/c, использование другого типа кабеля не доступно");
                        A.build[v].cabel_type = "оптоволокно";
                    }

                }

                Replase("ItogTrafik", Convert.ToString(load_num));

                Method.connect_load_calculation(ref M, A);
                M.connect_type = new string[M.connect_load.GetLength(0), M.connect_load.GetLength(0)];

                for (int v = 0; v < M.connect_load.GetLength(0); v++)
                    for (int f = v; f < M.connect_load.GetLength(0); f++)
                    {
                        bool check = false;
                        if (M.connect_load[v, f] != 0)
                        {

                            Console.WriteLine("Предполагаемая нагрузка между зданиями № " + (v + 1) + " и " + (f+1) + "= " + M.connect_load[v, f]+ " Расстояние = " +M.connect_length[v,f]);
                            if (M.connect_load[v, f] < 1000 && M.connect_length[v, f] < 100)
                            {
                                check = true;
                                Console.WriteLine("Для соединения между зданиями " + (v + 1) + " и " + (f + 1) + " подходит использование кабеля ethernet (0)");

                            }
                            if (M.connect_load[v, f] < 40000 && M.connect_length[v, f] < 3000)
                            {
                                check = true;
                                Console.WriteLine("Для соединения между зданиями " + (v + 1) + " и " + (f + 1) + " подходит использование оптоволоконного кабеля (1)");

                            }
                            if (M.connect_load[v, f] < 10 && M.connect_length[v, f] < 500)
                            {
                                check = true;
                                Console.WriteLine("Для соединения между зданиями " + (v + 1) + " и " + (f + 1) + " подходит использование коаксиального кабеля (2)");

                            }
                            if (!check)
                                Console.WriteLine("Необходимо использовать промежуточную инфраструктурy: \n 0 - ethernet, 1 - оптоволокно 2 - коаксиальный ");
                            
                            Console.WriteLine("Необходимо выбрать используемую технологию кабельного соединения ");
                            switch (Console.ReadLine())
                            {
                                case "0":
                                    {
                                        M.connect_type[v, f] = "ethernet";
                                        M.connect_type[f, v] = "ethernet";
                                        BetweenZdaniiCabel += "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель ethernet^l";
                                        itog_dlina_ethernet += M.connect_length[v, f];
                                    }
                                    break;
                                case "1":
                                    {
                                        M.connect_type[v, f] = "Оптоволокно";
                                        M.connect_type[f, v] = "Оптоволокно";
                                        itog_dlina_optovolokno += M.connect_length[v, f];
                                        BetweenZdaniiCabel += "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель кабеля^l";

                                    }
                                    break;
                                case "2":
                                    {
                                        M.connect_type[v, f] = "коаксиальный";
                                        M.connect_type[f, v] = "коаксиальный";
                                        itog_dlina_koaks += M.connect_length[v, f];
                                        BetweenZdaniiCabel += "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель кабеля^l";
                                    }
                                    break;
                            }
                        }
                    }
                //Console.ReadKey();
                Replase("@@CabelBetweenBuilds", BetweenZdaniiCabel);
                for (i = 0; i < Count; i++)

                {
                    int un_pow = 0;
                    Console.WriteLine("Здание № " + (i + 1));
                    //use_eq += "^l Для здания №" + (i + 1);
                    for (int i1 = 0; i1 < A.build[i].floor_num; i1++)
                    {

                        Console.WriteLine("Этаж № " + (i1 + 1));
                        Floor_req bl = Method.floor_req(ref A, i, i1);
                        itog_dlina_ethernet += bl.cabel_length;
                        cab_can_dlina_itog += bl.cab_can_length;
                        Console.WriteLine("Общая длина кабелей на этаже = " + bl.cabel_length);
                        CabelDlinaFloor += "Общая длина кабелей на этаже #" + Convert.ToString(i1 + 1) + " = " + Convert.ToString(bl.cabel_length) + "^l";
                        Console.WriteLine("Общая длина Кабель канала на этаже = " + bl.cab_can_length);
                        Console.WriteLine("Количество свитчей на этаже = " + bl.swch_num);
                        Console.WriteLine("Количество каналов свитчей на этаже = " + bl.swch_chan);
                        Console.WriteLine("Величина скорости свитчей на этаже = " + bl.swch_speed);
                        Console.WriteLine("Величина мощности питания = " + bl.power_req);
                        Console.WriteLine("Величина мощности бесперебойного питания = " + bl.power_uninter_req);
                        un_pow += bl.power_uninter_req;

                    }
                    use_eq += "Здание №" + Convert.ToString(i + 1) + "^l" + CabelDlinaFloor;
                    Console.WriteLine("Горизонтальная структура");
                    Build_req bl1 = Method.build_req(ref A, i);
                    switch (A.build[i].cabel_type)
                    {
                        case "оптоволокно":
                            itog_dlina_optovolokno += bl1.cabel_length;
                            break;
                        case "ethernet":
                            itog_dlina_ethernet += bl1.cabel_length;
                            break;

                    }
                    Console.WriteLine("Итоговая длина кабелей в здании № " + (i + 1) + " =" + bl1.cabel_length);
                    Console.WriteLine("Итоговая длина кабель-канала в здании № " + (i + 1) + " =" + bl1.cab_can_length);
                    Console.WriteLine("Итоговое количество свитчей в здании № " + (i + 1) + " =" + bl1.swch_num);
                    Console.WriteLine("Итоговое количество каналов свитчей в здании № " + (i + 1) + " =" + bl1.swch_chan);
                    Console.WriteLine("Итоговая размерность требуемой мощности бесперебойного питания в здании № " + (i + 1) + " =" + bl1.power_uninter_req);
                    cab_can_dlina_itog += bl1.cab_can_length;
                    uninterruptedpower += "^l Для здания " + i + " " + (un_pow + bl1.power_uninter_req);


                }
                Replase("CabelLengthFloor", use_eq);
                Replase("CabelLengthMax", "(" + (itog_dlina_ethernet + itog_dlina_koaks + itog_dlina_optovolokno) + "/305)=" + (Convert.ToInt32((itog_dlina_ethernet + itog_dlina_koaks + itog_dlina_optovolokno) / 305)));
                Console.WriteLine("\n\nИтоговая длина ethernet кабелей в системе =" + itog_dlina_ethernet);
                Console.WriteLine("Итоговая длина оптоволоконных кабелей в системе =" + itog_dlina_optovolokno);
                Console.WriteLine("Итоговая длина коаксиальных кабелей в системе =" + itog_dlina_koaks);
                Console.WriteLine("Итоговая длина кабель-каналов в системе =" + cab_can_dlina_itog);
                Replase("@@CabKan", cab_can_dlina_itog + "метров");
                Replase("@@uninterruptedpower", uninterruptedpower);
                Console.WriteLine("Нажмите Enter чтобы продолжить");
                Console.ReadLine();
                string use_tech = "";
                if (itog_dlina_ethernet != 0)
                    use_tech += "Gigabit Ethernet 1000 Base-TX, который основан на витой паре и волоконно-оптическом кабеле, ";
                if (itog_dlina_koaks != 0)
                    use_tech += "Gigabit Ethernet 10 Base-5, основанный на коаксиальном кабельном соединении, ";
                if (itog_dlina_optovolokno != 0)
                    use_tech += "Gigabit Ethernet 1000 Base-LX, использующий одномодовое волокно, ";
                IP[] ip =  Method.IP_ret(A);
                string[] ip_table_header = new string[3];
                ip_table_header[0] = "Номер здания";
                ip_table_header[1] = "Назначение";
                ip_table_header[2] = "Диапазон ip";
                string[,] ip_table = new string[ip.Length, 3]; 
                for (int z = 0; z < ip.Length; z++)
                {
                    ip_table[z, 0] = ip[z].build_num+"";
                    ip_table[z, 1] = ip[z].dep_name;
                    ip_table[z, 2] = ip[z].adress_start + "-" + ip[z].adress_end;
                }
                ReplaseTable("@@IP", ip_table, ip_table_header,7);
                Replase("@@techolog", use_tech);
                Replase("@@ equipment", use_eq);
                CloseDoc();
            //}
           /* catch
            {
                //Console.WriteLine("Произошла ошибка");
                //CloseDoc();
            }*/
          
            Console.Read();
        }
        static void Open_doc()
        {
            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = FilePatch; ;

            // если вылетим не этом этапе, приложение останется открытым
            try
            {
                document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                document.Close(ref falseObj, ref missingObj, ref missingObj);
                application.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                application = null;
                throw error;
            }
            application.Visible = true;

        }
        static void Replase(string strToFind, string replaceStr)
        {
            // обьектные строки для Word
            object strToFindObj = strToFind;
            object replaceStrObj = replaceStr;
            // диапазон документа Word
            Word.Range wordRange;
            //тип поиска и замены
            object replaceTypeObj;
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            int z = 0;
            while (replaceStr.Length > 0)
            // обходим все разделы документа
            {
                bool ch = false;
                if (255 - strToFind.Length > replaceStr.Length)
                    z = replaceStr.Length;
                else
                { z = 255 - strToFind.Length; ch = true; }
                string buf = string.Copy(replaceStr);

                if (ch)
                { buf += strToFind;buf= buf.Remove(z + 1); }
                    replaceStrObj = buf;
                for (int i = 1; i <= document.Sections.Count; i++)
                {

                    // берем всю секцию диапазоном
                    wordRange = document.Sections[i].Range;

                    Word.Find wordFindObj = wordRange.Find;

                    
                    object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };

                   
                    wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                }
                replaceStr = replaceStr.Remove(0, z);
            }
            Object pathToSaveObj = SaveFilePatch;
            document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj);

        }
        static void ReplaseTable(string strToFind, string[,] replaceStr, string[] headers,int tab_num)
        {
            // обьектные строки для Word
            object strToFindObj = strToFind;
            object replaceStrObj = replaceStr;
            // диапазон документа Word
            Word.Range wordRange;
            //тип поиска и замены
            object replaceTypeObj;
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            // обходим все разделы документа
            //for (int i = 1; i <= document.Sections.Count; i++)
            {

                // берем всю секцию диапазоном
                wordRange = document.Range();
                Word.Range buf_range;
                Word.Table table = document.Tables[tab_num];
                for (int rows = 0; rows < replaceStr.GetLength(0); rows++)
                    table.Rows.Add();
                for (int col = 0; col < replaceStr.GetLength(1); col++)
                {
                    for (int rows = 0; rows < replaceStr.GetLength(0); rows++)
                    {
                        buf_range = table.Cell(rows+2, col+1).Range;
                        string buf = replaceStr[rows, col];
                        buf_range.Text = buf;
                    }
                }

                /*
                Word.Find wordFindObj = wordRange.Find;
               // Word.Table table = document.Tables[tables_count];
               // tables_count++;
                Word.Range buf_range;
                Word.Table table = new Word.Table;
                for(int col=0;i< replaceStr.GetLength(0); i++)
                    table.Columns.Add();
                for (int rows = 0; i < replaceStr.GetLength(1)+1; i++)
                    table.Rows.Add();
                for (int col = 0; i < replaceStr.GetLength(0); i++)
                {
                    buf_range = table.Cell(col, 1).Range;
                    buf_range.Text = headers[col];
                    for (int rows = 1; i < replaceStr.GetLength(1); i++)
                    {
                        buf_range = table.Cell(col, rows).Range;
                        buf_range.Text = replaceStr[rows,col];
                    }
                }
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, table, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
*/
            }
            Object pathToSaveObj = SaveFilePatch;
            document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocument, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj);

        }

        static void CloseDoc()
        {
            document.Close(ref falseObj, ref missingObj, ref missingObj);
            application.Quit(ref missingObj, ref missingObj, ref missingObj);
        }


    }

}