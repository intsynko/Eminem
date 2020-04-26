using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Emionov_root;

namespace Eminem
{
    class Program
    {
        static string CabelDlinaFloor;
        static string BetweenZdaniiCabel = "Нет необходимости соединения зданий";

        [STAThreadAttribute] // тэг нужен для работы диалогово окна
        static void Main(string[] args)
        {
            float itog_dlina_ethernet = 0;
            float itog_dlina_optovolokno = 0;
            int itog_dlina_koaks = 0;
            int cab_can_dlina_itog = 0;
            bool release = true;
#if DEBUG
    release = false;
#endif

            // контекстный менеджер, правильно закроет файл, если внутри блока произойдет какая нибудь фигня (исключение)
            using (MyDocument document = new MyDocument(killAllProcesses: true, askByDialogWindow: release) {Visible = false})
            {
                string message = "Введите номер вашего варианта, первая цифра это последняя цифра номера группы, вторая и третья цифра это ваш вариант.\n Например, 901: ";
                int nom_var = GetInt(message);
                document.Replase("VariantKurs", nom_var.ToString());

                message = "Введите количество используемых Информационных систем: ";
                int InformationSystemCount = GetInt(message);
                Technology[] TechnologyArray = new Technology[InformationSystemCount];
                string[] tehnol_tab_header = new string[] {
                    "Решаемая задача",
                    "Информационная система", 
                    "Характер решаемой задачи"
                };
                
                // заполнение сведений о ИС (, название,)
                string[,] tehno_tab = new string[InformationSystemCount, 3];
                for (int k = 0; k < InformationSystemCount; k++)
                {
                    Console.Write($"Введите название Информационной системы № {k+1}: ");
                    TechnologyArray[k].name = Console.ReadLine();
                    tehno_tab[k, 1] = TechnologyArray[k].name;
                    message = $"Введите количество трафика Информационной системы в мб/с от пользователя (только целое число) {TechnologyArray[k].name}: ";
                    TechnologyArray[k].load = GetInt(message);

                }
                document.ReplaseTable("@@system_table", tehno_tab, tehnol_tab_header, 3);

                Builds BuildsObj = new Builds();
                int BuildsCount = GetInt("Введите количество зданий: ");
                document.Replase("Kolzdanii", BuildsCount.ToString());
                

                int[,] matrix = new int[BuildsCount, BuildsCount];
                string betw_buil = "-";
                // если имеется больше одного здания
                if (BuildsCount > 1)
                {
                    // создаем связки зданий
                    int pari = GetInt("Сколько существует пар зданий связанных: ");
                    for (int z = 0; z < pari; z++)
                    {
                        int first_build = GetInt("Введите номер первого из двух связанных зданий: ", matrix.GetLength(0));
                        int second_build = GetInt("Введите номер второго из двух связанных зданий: ", matrix.GetLength(0));
                        int length = GetInt("Введите расстояние между этими 2 зданиями: ");
                        matrix[first_build - 1, second_build - 1] = length;
                        matrix[second_build - 1, first_build - 1] = length;
                        betw_buil += first_build + " - " + second_build + " = " + length + ", ";
                    }
                }

                document.Replase("BetweenBuilds", betw_buil);

                string[] depart_tech_haeders = new string []{
                    "Название отдела",
                    "Способ взаимодействия",
                    "Информационная система",
                    "Организационные единицы",
                    "Характер решаемой задачи"
                };
                int department_counter = 0;
                string[,] depart_tech = new string[department_counter + 1, 5];
                Connect ConnectObj = new Connect();
                ConnectObj.connect_length = matrix;
                int i, all_workers = 0;
                BuildsObj.build = new Build[BuildsCount];
                // божественноое объявление преременных
                string floor_num = "", squre_num = "", hight_num = "", workers_num = "", mob_st_num = "", descr_otd = "", use_eq = "", uninterruptedpower = "";
                for (i = 0; i < BuildsCount; i++)
                {
                    Console.WriteLine($"----------Здание {i + 1}----------------");
                    descr_otd += $"В здании №{i+1} находятся: ";
                    BuildsObj.build[i] = new Build();
                    BuildsObj.build[i].floor_num = GetInt($"Введите количество этажей здания №{i + 1}: ");
                    floor_num += BuildsObj.build[i].floor_num + ",";
                    BuildsObj.build[i].square = GetInt($"Введите площадь этажа здания №{i + 1}: ");
                    squre_num += BuildsObj.build[i].square + ",";
                    BuildsObj.build[i].height = GetFlaot($"Введите высоту этажа здания №{i + 1}");
                    hight_num += BuildsObj.build[i].height;
                    int mob_st = GetInt($"Введите количество мобильных станций здания №{i + 1}: ");
                    mob_st_num += mob_st + ",";
                    workers_num += i + " ";

                    BuildsObj.build[i].floor = new Floor[BuildsObj.build[i].floor_num];
                    for (int i1 = 0; i1 < BuildsObj.build[i].floor_num; i1++)
                    {
                        Console.WriteLine($"----------Этаж {i1 + 1}----------------");
                        descr_otd += $"{i1+1} Этаж: ";
                        int workers = 0;

                        Console.WriteLine("----------Напоминалка------------");
                        Console.WriteLine("Напоминание, под каким номером какая Информационная система: ");
                        for (int k = 0; k < InformationSystemCount; k++)
                            Console.WriteLine($"Информационная система № {k + 1}: {TechnologyArray[k].name}");
                        Console.WriteLine("--------------------------------");

                        BuildsObj.build[i].floor[i1] = new Floor();
                        BuildsObj.build[i].floor[i1].Mob_st = mob_st;
                        BuildsObj.build[i].floor[i1].dep_num = GetInt($"Введите количество отделов этажа № {i1 + 1}  Здания № {i + 1}: ");
                        BuildsObj.build[i].floor[i1].dep = new Department[BuildsObj.build[i].floor[i1].dep_num];
                        for (int i2 = 0; i2 < BuildsObj.build[i].floor[i1].dep_num; i2++)
                        {
                            Console.WriteLine("Напоминание, под каким номером какая Информационная система:");
                            for (int k = 0; k < InformationSystemCount; k++)
                                Console.WriteLine($"Информационная система № {k + 1}: {TechnologyArray[k].name}");

                            Console.WriteLine($"----------Отдел {i2 + 1}----------------");
                            Console.Write($"Введите название отдела: ");
                            BuildsObj.build[i].floor[i1].dep[i2].name = Console.ReadLine();
                            depart_tech[department_counter, 0] = BuildsObj.build[i].floor[i1].dep[i2].name;

                            message = $"Введите количество работников отдела:" ;
                            BuildsObj.build[i].floor[i1].dep[i2].workers = GetInt(message);
                            workers += BuildsObj.build[i].floor[i1].dep[i2].workers;
                            all_workers += BuildsObj.build[i].floor[i1].dep[i2].workers;

                            message = $"Введите количество используемых Информационных систем отдела: ";
                            BuildsObj.build[i].floor[i1].dep[i2].techno_count = GetInt(message);
                            BuildsObj.build[i].floor[i1].dep[i2].tech = new Technology[BuildsObj.build[i].floor[i1].dep[i2].techno_count];
                            for (int i3 = 0; i3 < BuildsObj.build[i].floor[i1].dep[i2].techno_count; i3++)
                            {

                                message = "Введите номер Информационной системы, которая используется в отделе: ";
                                BuildsObj.build[i].floor[i1].dep[i2].tech[i3] = TechnologyArray[GetInt(message, TechnologyArray.Length) - 1];
                                depart_tech[department_counter, 2] += BuildsObj.build[i].floor[i1].dep[i2].tech[i3].name + "^P"; // что это за знак???
                                BuildsObj.build[i].floor[i1].dep[i2].tech[i3].user = true;
                                depart_tech[department_counter, 1] += "^i";
                                depart_tech[department_counter, 1] += "Пользователь";

                                if (Question("Является ли отдел корневым отделом, к которому идут все запросы?"))
                                {
                                    BuildsObj.build[i].floor[i1].dep[i2].tech[i3].root = true;
                                    depart_tech[department_counter, 1] += ",Сервер";
                                }
                                else
                                    BuildsObj.build[i].floor[i1].dep[i2].tech[i3].root = false;

                                BuildsObj.build[i].floor[i1].dep[i2].tech[i3].rem_serv = Question(
                                    "Использует ли программа ресурсы сети интернет, " +
                                    "если да - все запросы считаются внесетевыми и не рассчитываются " +
                                    "между зданиями."
                                );
                            }
                            Department dep = BuildsObj.build[i].floor[i1].dep[i2];
                            descr_otd += $"{dep.name} ({dep.workers} рабочих станций), ";
                            department_counter++;
                            string[,] buf_mas = new string[department_counter + 1, 5];
                            Array.Copy(depart_tech, buf_mas, depart_tech.Length);
                            depart_tech = buf_mas;
                            Console.WriteLine("--------------------------------");

                        }
                        workers_num += workers + ",";
                        Console.WriteLine("--------------------------------");
                    }
                    workers_num += "^l";
                    Console.WriteLine("--------------------------------");
                }
                Console.WriteLine("Подождите, идет заполнение документа");
                document.ReplaseTable("@@depart_tech", depart_tech, depart_tech_haeders, 6);
                //вывод в таблицу
                document.Replase("FloorsBuilds", floor_num);
                document.Replase("FloorSquare", squre_num);
                document.Replase("FloorHeight", hight_num);
                document.Replase("WorkersFloors", workers_num);
                document.Replase("MobStations", mob_st_num);
                Console.WriteLine("Введите время реакции системы в мс. Например, 400:");
                string rofel7 = Console.ReadLine();
                document.Replase("ReactionTime", rofel7);
                document.Replase("WorkersBuilds", all_workers + " ");

                document.Replase("@@Otdel", descr_otd);// заполняем описания отделов
                int load_num = 0;

                Method.build_load_calculation(ref BuildsObj);
                for (int v = 0; v < BuildsObj.load.Length; v++)
                {
                    load_num += BuildsObj.load[v];
                    Console.WriteLine("Предполагаемая нагрузка внутри здания № " + (v + 1) + " =" + BuildsObj.load[v]);
                    if (BuildsObj.load[v] < 1000)
                    {
                        BuildsObj.build[v].cabel_type = "ethernet";
                        Console.WriteLine("В здании № " + (v + 1) + " предполагается использование кабеля ethernet");

                        if (BuildsObj.build[v].height * BuildsObj.build[v].floor_num > 100)
                            Console.WriteLine("Требуется применение дополнительных устройств усиления для вертикальных кабелей");

                        float len = 0;
                        for (int i5 = 0; i5 < BuildsObj.build[v].floor_num; i5++)
                            len += BuildsObj.build[v].height * i5;

                        if (!Question(""))
                        {
                            BuildsObj.build[v].cabel_type = "оптоволокно";
                            itog_dlina_optovolokno = len;
                        }
                        else
                            itog_dlina_ethernet += len;
                    }
                    else
                    {
                        Console.WriteLine(
                            $"В здании № { (v + 1)} предполагается использование оптоволоконного кабеля," +
                            $" поскольку нагрузка на сеть больше 1000 Мб/c," +
                            $" использование другого типа кабеля не доступно");
                        BuildsObj.build[v].cabel_type = "оптоволокно";
                    }

                }

                document.Replase("ItogTrafik", Convert.ToString(load_num));

                Method.connect_load_calculation(ref ConnectObj, BuildsObj);
                ConnectObj.connect_type = new string[ConnectObj.connect_load.GetLength(0), ConnectObj.connect_load.GetLength(0)];

                for (int v = 0; v < ConnectObj.connect_load.GetLength(0); v++)
                {
                    for (int f = v; f < ConnectObj.connect_load.GetLength(0); f++)
                    {
                        bool check = false;
                        if (ConnectObj.connect_load[v, f] != 0)
                        {

                            Console.WriteLine("Предполагаемая нагрузка между зданиями № " + (v + 1) + " и " + (f + 1) + "= " + ConnectObj.connect_load[v, f] + " Расстояние = " + ConnectObj.connect_length[v, f]);
                            if (ConnectObj.connect_load[v, f] < 1000 && ConnectObj.connect_length[v, f] < 100)
                            {
                                check = true;
                                Console.WriteLine("Для соединения между зданиями " + (v + 1) + " и " + (f + 1) + " подходит использование кабеля ethernet (0)");

                            }
                            if (ConnectObj.connect_load[v, f] < 40000 && ConnectObj.connect_length[v, f] < 3000)
                            {
                                check = true;
                                Console.WriteLine("Для соединения между зданиями " + (v + 1) + " и " + (f + 1) + " подходит использование оптоволоконного кабеля (1)");

                            }
                            if (ConnectObj.connect_load[v, f] < 10 && ConnectObj.connect_length[v, f] < 500)
                            {
                                check = true;
                                Console.WriteLine("Для соединения между зданиями " + (v + 1) + " и " + (f + 1) + " подходит использование коаксиального кабеля (2)");

                            }
                            if (!check)
                                Console.WriteLine("Необходимо использовать промежуточную инфраструктурy: \n 0 - ethernet, 1 - оптоволокно 2 - коаксиальный ");

                            Console.WriteLine("Необходимо выбрать используемую технологию кабельного соединения: ");
                            switch (Console.ReadLine())
                            {
                                case "0":
                                    {
                                        ConnectObj.connect_type[v, f] = "ethernet";
                                        ConnectObj.connect_type[f, v] = "ethernet";
                                        BetweenZdaniiCabel = "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель ethernet^l";
                                        itog_dlina_ethernet += ConnectObj.connect_length[v, f];
                                    }
                                    break;
                                case "1":
                                    {
                                        ConnectObj.connect_type[v, f] = "Оптоволокно";
                                        ConnectObj.connect_type[f, v] = "Оптоволокно";
                                        itog_dlina_optovolokno += ConnectObj.connect_length[v, f];
                                        BetweenZdaniiCabel = "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель кабеля^l";

                                    }
                                    break;
                                case "2":
                                    {
                                        ConnectObj.connect_type[v, f] = "коаксиальный";
                                        ConnectObj.connect_type[f, v] = "коаксиальный";
                                        itog_dlina_koaks += ConnectObj.connect_length[v, f];
                                        BetweenZdaniiCabel = "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель кабеля^l";
                                    }
                                    break;
                            }
                            document.Replase("@@CabelBetweenBuilds", BetweenZdaniiCabel);
                        }
                    }
                }
                //Console.ReadKey();
                
                for (i = 0; i < BuildsCount; i++)

                {
                    int un_pow = 0;
                    Console.WriteLine("Здание № " + (i + 1));
                    //use_eq += "^l Для здания №" + (i + 1);
                    for (int i1 = 0; i1 < BuildsObj.build[i].floor_num; i1++)
                    {

                        Console.WriteLine("Этаж № " + (i1 + 1));
                        Floor_req bl = Method.floor_req(ref BuildsObj, i, i1);
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
                    Build_req bl1 = Method.build_req(ref BuildsObj, i);
                    switch (BuildsObj.build[i].cabel_type)
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
                document.Replase("CabelLengthFloor", use_eq);
                document.Replase("CabelLengthMax", "(" + (itog_dlina_ethernet + itog_dlina_koaks + itog_dlina_optovolokno) + "/305)=" + (int)((itog_dlina_ethernet + itog_dlina_koaks + itog_dlina_optovolokno) / 305));
                Console.WriteLine("\n\nИтоговая длина ethernet кабелей в системе =" + itog_dlina_ethernet);
                Console.WriteLine("Итоговая длина оптоволоконных кабелей в системе =" + itog_dlina_optovolokno);
                Console.WriteLine("Итоговая длина коаксиальных кабелей в системе =" + itog_dlina_koaks);
                Console.WriteLine("Итоговая длина кабель-каналов в системе =" + cab_can_dlina_itog);
                document.Replase("@@CabKan", cab_can_dlina_itog + "метров");
                document.Replase("@@uninterruptedpower", uninterruptedpower);
                Console.WriteLine("Нажмите Enter чтобы продолжить");
                Console.ReadLine();
                string use_tech = "";
                if (itog_dlina_ethernet != 0)
                    use_tech += "Gigabit Ethernet 1000 Base-TX, который основан на витой паре и волоконно-оптическом кабеле, ";
                if (itog_dlina_koaks != 0)
                    use_tech += "Gigabit Ethernet 10 Base-5, основанный на коаксиальном кабельном соединении, ";
                if (itog_dlina_optovolokno != 0)
                    use_tech += "Gigabit Ethernet 1000 Base-LX, использующий одномодовое волокно, ";
                IP[] ip = Method.IP_ret(ref BuildsObj);
                string[] ip_table_header = new string[3];
                ip_table_header[0] = "Номер здания";
                ip_table_header[1] = "Назначение";
                ip_table_header[2] = "Диапазон ip";
                string[,] ip_table = new string[ip.Length, 3];
                for (int z = 0; z < ip.Length; z++)
                {
                    ip_table[z, 0] = ip[z].build_num + "";
                    ip_table[z, 1] = ip[z].dep_name;
                    ip_table[z, 2] = ip[z].adress_start + "-" + ip[z].adress_end;
                }
                document.ReplaseTable("@@IP", ip_table, ip_table_header, 7);
                document.Replase("@@techolog", use_tech);
                document.Replase("@@ equipment", use_eq);
                Console.Read();
            }
            
        }



        /// <summary>
        /// Запрошивает у пользователя число в формате float, пока не введет в верном формате 
        /// </summary>
        /// <param name="message">Сообщеине, которое спросят у пользователя в коносли</param>
        /// <param name="checkNegativeAndkZero"></param>
        /// <returns>Число</returns>
        static float GetFlaot(string message, bool checkNegativeAndkZero = true)
        {
            while (true)
            {
                try
                {
                    Console.Write($"{message} (дробное число пешется через запятую, например 3,5):");
                    float num = (float)Convert.ToDouble(Console.ReadLine());
                    if (checkNegativeAndkZero && num <= 0) throw new OverflowException("Число меньше или равно нулю.");
                    return num;
                }
                catch (FormatException ex)
                {
                    Console.WriteLine("Вы некорректно ввели число.");
                }
                catch (OverflowException ex)
                {
                    Console.WriteLine("Число выходит за допустимые пределы.");
                }
            }

        }


        /// <summary>
        /// Запрошивает у пользователя число в формате int, пока не введет в верном формате
        /// </summary>
        /// <param name="message">Сообщеине, которое спросят у пользователя в коносли</param>
        /// <param name="checkNegativeAndkZero">Проверять на отриц./нулевое значение</param>
        /// <returns>Число</returns>
        static int GetInt(string message, bool checkNegativeAndkZero = true)
        {   
            while(true)
            {
                try
                {
                    Console.Write(message);
                    int num = Convert.ToInt32(Console.ReadLine());
                    if (checkNegativeAndkZero && num <= 0) throw new OverflowException("Число меньше или равно нулю.");
                    return num;
                }
                catch (FormatException ex){
                    Console.WriteLine("Вы некорректно ввели число.");
                }
                catch (OverflowException ex){
                    Console.WriteLine("Число выходит за допустимые пределы.");
                }
            }
            
        }
        /// <summary>
        /// Запрошивает у пользователя число в формате int, пока не введет в верном формате
        /// </summary>
        /// <param name="message">Сообщеине, которое спросят у пользователя в коносли</param>
        /// <param name="max">Верхняя допустимая граница (включительно)</param>
        /// <param name="min">Нижняя граница (не включительно)</param>
        /// <returns></returns>
        static int GetInt(string message, int max, int min = 0)
        {
            while (true)
            {
                int num = GetInt(message);
                if (num > min && num <= max)
                    return num;
                else
                    Console.WriteLine(
                        "Вы ввели число, выходящее за границы массива, \n" +
                        $"напоминаем, что граница массива: 1...{max}"
                    );
            }

        }

        /// <summary>
        /// Запрошивает у вопрос с ответами да и нет
        /// </summary>
        /// <param name="message">Сообщение вопроса</param>
        /// <returns></returns>
        static bool Question(string message)
        {
            Console.WriteLine($"{message} \n Да - '1' Нет - '2': ");
            string qw;
            do { qw = Console.ReadLine(); } while (qw != "1" && qw != "2");
            return qw == "1";
        }
    }

}