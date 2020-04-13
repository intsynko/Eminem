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


        static void Main(string[] args)
        {
            int itog_dlina_ethernet = 0;
            int itog_dlina_optovolokno = 0;
            int itog_dlina_koaks = 0;
            int cab_can_dlina_itog = 0;

            Console.Write("Осторожно! Убивает все открытые процессы ворда при запуске.\nДля продолжения нажмите ENTER:");
            Console.ReadLine();

            // контекстный менеджер, правильно закроет файл, если во время блока произойдет какая нибудь фигня (исключение)
            using (MyDocument document = new MyDocument(true) {Visible = false})
            {
                //ToDo поменять все читания инта с консоли на метод GetInt
                //ToDo поменять все каскады Console.WriteLine(...) на интерполяию
                //ToDo вообще применить интерполяцию там, где это возможно

                string message = "Введите номер вашего варианта, первая цифра это последняя цифра номера группы, вторая и третья цифра это ваш вариант.\n Например, 901: ";
                int nom_var = GetInt(message);
                document.Replase("VariantKurs", nom_var.ToString());

                message = "Введите количество используемых Информационных систем: ";
                int Tech_number = GetInt(message);
                Technology[] T = new Technology[Tech_number];
                string[] tehnol_tab_header = new string[3];
                tehnol_tab_header[0] = "Решаемая задача";
                tehnol_tab_header[1] = "Информационная система";
                tehnol_tab_header[2] = "Характер решаемой задачи";
                string[,] tehno_tab = new string[Tech_number, 3];
                for (int k = 0; k < Tech_number; k++)
                {

                    Console.Write($"Введите название Информационной системы № {k+1}: ");
                    T[k].name = Console.ReadLine();
                    tehno_tab[k, 1] = T[k].name;
                    message = $"Введите количество трафика Информационной системы в мб/с от пользователя (только целое число) {T[k].name}: ";
                    T[k].load = GetInt(message);

                }
                document.ReplaseTable("@@system_table", tehno_tab, tehnol_tab_header, 3);

                Builds A = new Builds();
                int Count = GetInt("Введите количество зданий: ");
                document.Replase("Kolzdanii", Count.ToString());
                /*
               Console.WriteLine("Введите расстояния между зданиями, через запятую. Например: 400,600:");
               string rofel = Console.ReadLine();
               document.Replase("BetweenBuilds", rofel);
               Console.WriteLine("Введите количество этажей в зданиях, через запятую. Например: 4,3:");
               string rofel1 = Console.ReadLine();
               document.Replase("FloorBuilds", rofel1);
               Console.WriteLine("Введите площадь этажей в зданиях, через запятую. Например: 400,300:");
               string rofel2 = Console.ReadLine();
               document.Replase("FloorSquare", rofel2);
               Console.WriteLine("Введите высоту этажей в зданиях, через запятую. Например: 4,3:");
               string rofel3 = Console.ReadLine();
               document.Replase("FloorHeight", rofel3);

               Console.WriteLine("Введите количество рабочих по этажам, через запятую. Например: 400 на первом,300 на втором,200 на остальных:");
               string rofel5 = Console.ReadLine();
               document.Replase("WorkersFloors", rofel5);
               Console.WriteLine("Введите количество мобильных станций в помещении. Например, 40:");
               string rofel6 = Console.ReadLine();
               document.Replase("MobStations", rofel6);
               */

                /*

               */

                int[,] matrix = new int[Count, Count];
                string betw_buil = "-";
                // если имеется больше одного здания
                if (Count > 1)
                {
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
                int i, all_workers = 0;
                A.build = new Build[Count];
                string floor_num = "", squre_num = "", hight_num = "", workers_num = "", mob_st_num = "", descr_otd = "", use_eq = "", uninterruptedpower = "";
                for (i = 0; i < Count; i++)
                {
                    Console.WriteLine($"----------Здание {i + 1}----------------");
                    descr_otd += " В здании №" + i + "находятся: ";
                    A.build[i] = new Build();
                    A.build[i].floor_num = GetInt($"Введите количество этажей здания №{i + 1}: ");
                    floor_num += A.build[i].floor_num + ",";
                    A.build[i].square = GetInt($"Введите площадь этажа здания №{i + 1}: ");
                    squre_num += A.build[i].square + ",";
                    A.build[i].height = GetInt($"Введите высоту этажа здания №{i + 1}: ");
                    hight_num += A.build[i].height;
                    int mob_st = GetInt($"Введите количество мобильных станций здания №{i + 1}: ");
                    mob_st_num += mob_st + ",";
                    workers_num += i + " ";

                    A.build[i].floor = new Floor[A.build[i].floor_num];
                    for (int i1 = 0; i1 < A.build[i].floor_num; i1++)
                    {
                        Console.WriteLine($"----------Этаж {i1 + 1}----------------");
                        descr_otd += i1 + "этаж - ";
                        int workers = 0;
                        A.build[i].floor[i1] = new Floor();
                        A.build[i].floor[i1].Mob_st = mob_st;
                        A.build[i].floor[i1].dep_num = GetInt($"Введите количество отделов этажа: ");
                        A.build[i].floor[i1].dep = new Department[A.build[i].floor[i1].dep_num];

                        Console.WriteLine("----------Напоминалка------------");
                        Console.WriteLine("Напоминание, под каким номером какая Информационная система: ");
                        for (int k = 0; k < Tech_number; k++)
                            Console.WriteLine($"Информационная система № {k + 1}: {T[k].name}");
                        Console.WriteLine("--------------------------------");

                        for (int i2 = 0; i2 < A.build[i].floor[i1].dep_num; i2++)
                        {
                            Console.WriteLine($"----------Отдел {i2 + 1}------------");
                            Console.Write($"Введите название отдела № {i2 + 1}  этажа № {i + 1}: ");
                            A.build[i].floor[i1].dep[i2].name = Console.ReadLine();
                            depart_tech[department_count, 0] = A.build[i].floor[i1].dep[i2].name;
                            message = $"Введите количество работников отдела:" ;
                            A.build[i].floor[i1].dep[i2].workers = GetInt(message);
                            workers += A.build[i].floor[i1].dep[i2].workers;
                            all_workers += A.build[i].floor[i1].dep[i2].workers;
                            message = $"Введите количество используемых Информационных систем отдела: ";
                            A.build[i].floor[i1].dep[i2].techno_count = GetInt(message);
                            A.build[i].floor[i1].dep[i2].tech = new Technology[A.build[i].floor[i1].dep[i2].techno_count];
                            for (int i3 = 0; i3 < A.build[i].floor[i1].dep[i2].techno_count; i3++)
                            {

                                message = "Введите номер Информационной системы, которая используется в отделе: ";

                                A.build[i].floor[i1].dep[i2].tech[i3] = T[GetInt(message, T.Length) - 1];
                                depart_tech[department_count, 2] += A.build[i].floor[i1].dep[i2].tech[i3].name + "^P"; // что это за знак???
                                A.build[i].floor[i1].dep[i2].tech[i3].user = true;
                                depart_tech[department_count, 1] += "Пользователь";

                                string qw = "";
                                Console.WriteLine("Является ли отдел корневым отделом," +
                                    " к которому идут все запросы? Да - '1' Нет - '2': ");
                                do { qw = Console.ReadLine();} while (qw != "1" || qw != "2");
                                if (qw == "1")
                                {
                                    A.build[i].floor[i1].dep[i2].tech[i3].root = true;
                                    depart_tech[department_count, 1] += ",Сервер";
                                }
                                else
                                    A.build[i].floor[i1].dep[i2].tech[i3].root = false;

                                depart_tech[department_count, 1] += "^i";
                                Console.WriteLine("Использует ли программа ресурсы сети интернет, " +
                                    "если да - все запросы считаются внесетевыми и не рассчитываются " +
                                    "между зданиями. Да - '1' Нет - '2': ");

                                do { qw = Console.ReadLine(); } while (qw != "1" || qw != "2");

                                if (qw == "1")
                                    A.build[i].floor[i1].dep[i2].tech[i3].rem_serv = true;
                                else
                                    A.build[i].floor[i1].dep[i2].tech[i3].rem_serv = false;
                            }
                            descr_otd += A.build[i].floor[i1].dep[i2].name + " (" + A.build[i].floor[i1].dep[i2].workers + " рабочих станций), ";
                            department_count++;
                            string[,] buf_mas = new string[department_count + 1, 5];
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

                Method.build_load_calculation(ref A);
                for (int v = 0; v < A.load.Length; v++)
                {
                    load_num += A.load[v];
                    Console.WriteLine("Предполагаемая нагрузка внутри здания № " + (v + 1) + " =" + A.load[v]);
                    if (A.load[v] < 1000)
                    {
                        A.build[v].cabel_type = "ethernet";
                        Console.WriteLine("В здании № " + (v + 1) + " предполагается использование кабеля ethernet");

                        if (A.build[v].height * A.build[v].floor_num > 100)
                            Console.WriteLine("Требуется применение дополнительных устройств усиления для вертикальных кабелей");
                        Console.WriteLine("Вы согласны? 1 - Да, 2 - Нет ");

                        int len = 0;
                        for (int i5 = 0; i5 < A.build[v].floor_num; i5++)
                            len += A.build[v].height * i5;

                        if (Console.ReadLine() == "2")
                        {
                            A.build[v].cabel_type = "оптоволокно";
                            itog_dlina_optovolokno = len;
                        }
                        else
                            itog_dlina_ethernet += len;
                    }
                    else
                    {
                        Console.WriteLine(
                            "В здании № " + (v + 1) + " предполагается использование " +
                            "оптоволоконного кабеля, поскольку нагрузка на сеть больше 1000 Мб/c, " +
                            "использование другого типа кабеля не доступно"
                        );
                        A.build[v].cabel_type = "оптоволокно";
                    }

                }

                document.Replase("ItogTrafik", Convert.ToString(load_num));

                Method.connect_load_calculation(ref M, A);
                M.connect_type = new string[M.connect_load.GetLength(0), M.connect_load.GetLength(0)];

                for (int v = 0; v < M.connect_load.GetLength(0); v++)
                    for (int f = v; f < M.connect_load.GetLength(0); f++)
                    {
                        bool check = false;
                        if (M.connect_load[v, f] != 0)
                        {

                            Console.WriteLine("Предполагаемая нагрузка между зданиями № " + (v + 1) + " и " + (f + 1) + "= " + M.connect_load[v, f] + " Расстояние = " + M.connect_length[v, f]);
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

                            Console.WriteLine("Необходимо выбрать используемую технологию кабельного соединения: ");
                            switch (Console.ReadLine())
                            {
                                case "0":
                                    {
                                        M.connect_type[v, f] = "ethernet";
                                        M.connect_type[f, v] = "ethernet";
                                        BetweenZdaniiCabel = "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель ethernet^l";
                                        itog_dlina_ethernet += M.connect_length[v, f];
                                    }
                                    break;
                                case "1":
                                    {
                                        M.connect_type[v, f] = "Оптоволокно";
                                        M.connect_type[f, v] = "Оптоволокно";
                                        itog_dlina_optovolokno += M.connect_length[v, f];
                                        BetweenZdaniiCabel = "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель кабеля^l";

                                    }
                                    break;
                                case "2":
                                    {
                                        M.connect_type[v, f] = "коаксиальный";
                                        M.connect_type[f, v] = "коаксиальный";
                                        itog_dlina_koaks += M.connect_length[v, f];
                                        BetweenZdaniiCabel = "Для соединения между зданиями " + Convert.ToString(v) + " и " + Convert.ToString(f) + "будет использован кабель кабеля^l";
                                    }
                                    break;
                            }
                            document.Replase("@@CabelBetweenBuilds", BetweenZdaniiCabel);
                        }
                    }
                //Console.ReadKey();
                
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
                IP[] ip = Method.IP_ret(A);
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
            }
            Console.Read();
        }


        
        /// <summary>
        /// Запрошивает у пользователя число в формате int, пока не введет в верном формате
        /// </summary>
        /// <param name="message">Сообщеине, которое спросят у пользователя в коносли</param>
        /// <returns>Число</returns>
        static int GetInt(string message, bool checkZero = true)
        {   
            while(true)
            {
                try
                {
                    Console.Write(message);
                    int num = Convert.ToInt32(Console.ReadLine());
                    if (checkZero && num <= 0) throw new OverflowException("Число меньше или равно нулю.");
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
    }

}