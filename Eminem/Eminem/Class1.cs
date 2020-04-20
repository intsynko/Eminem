using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;



namespace Emionov_root
{
    
    //
    //Работа со зданиями
    //
    public struct Builds
    {
        public Build[] build;
        public int[] load;
        
    }
    public struct IP
    {
        public int build_num;
        public string dep_name;
        public string adress_start;
        public string adress_end;
    }

    public struct Build //описание здания
    {
        public string cabel_type;
        public int floor_num;
        public int square;
        public float height;
        public Floor[] floor; //создаем массив с кол-вом этажей и их перечислениeм в каждой ячейке
        public Build_req build_req;

    }
    public struct Floor
    {
        public int dep_num;
        public Department[] dep; //создаем массив с кол-вом отделов на этаже и их перечисленим в каждой ячейке
        public int Mob_st;
        public Floor_req fl_req;
    }
    public struct Department
    {
        public String name;
        public int workers;
        public int techno_count;
        public Technology[] tech; //список технологий, используемых в отделе
    }
    public struct Technology
    {
        public string name;
        public int load; // нагрузка в Мб!!!!
        public bool root; //является ли отдел корневым отделом, к которому идут все запросы
        public bool user; //является ли отдел пользователем технологии
        public bool rem_serv; //использует ли программа ресурсы сети интернет, если да - все запросы считаются внесетевыми и не рассчитываются между зданиями
    }
    public struct Floor_req //только для расчета на этаже NADO VIVESTI
    {
        public int cabel_length;
        public int cab_can_length;
        public int swch_num;
        public int swch_chan;
        public int swch_speed;
        public int power_req;
        public int power_uninter_req;
    }
    public struct Build_req //только для расчета между этажами NADO VIVESTI
    {
        public float cabel_length;
        public int cab_can_length;
        public int swch_num; //возможно расположение некоторых свичей на разных этажах, при 2+ свичах третий с магистрального допом
        public int swch_chan;
        public int power_uninter_req; //только для главных серверных
    }
    public struct Connect
    {
        public int[,] connect_length; //дб КВАДРАТНЫМ!!!
        public int[,] connect_load;
        public string[,] connect_type;
    } 

    public class Method
    {
        static int swch_pow = 50;
        static int serv_pow = 1000;
        static int pc_pow = 500;
        static int swch_ch_max = 50;


        //fie@list.ru



        public static void build_load_calculation(ref Builds builds)
        {
            // ToDo переделать все это на методы классов

            // создаем массив с нагрузками на здания
            builds.load = new int[builds.build.Length]; 
            for (int i=0; i < builds.build.Length; i++)
            {
                for (int fl=0; fl <  builds.build[i].floor.Length; fl++)
                {
                    for (int de = 0; de < builds.build[i].floor[fl].dep.Length; de++)
                    {
                        for (int te=0; te < builds.build[i].floor[fl].dep[de].tech.Length; te++)
                        {
                            // если это корневой отдел
                            if (builds.build[i].floor[fl].dep[de].tech[te].root)
                            {
                                // считаем всех пользователей этой ИС
                                string name_to_find = builds.build[i].floor[fl].dep[de].tech[te].name;
                                int client = builds.build.Sum(
                                    b => b.floor.Sum(
                                        f => f.dep.Where(
                                            d => d.tech.Select(t => t.name).Contains(name_to_find)
                                        ).Sum(d => d.workers)
                                    )
                                );
                                //считаю нагрузку на сервера
                                builds.load[i] += client * builds.build[i].floor[fl].dep[de].tech[te].load;
                                builds.build[i].floor[fl].fl_req.swch_speed+= client * builds.build[i].floor[fl].dep[de].tech[te].load;
                            }
                            // если отдел использует эту ИС
                            if(builds.build[i].floor[fl].dep[de].tech[te].user)
                            {
                                //считаю нагрузку клиентов
                                builds.load[i] += builds.build[i].floor[fl].dep[de].workers * builds.build[i].floor[fl].dep[de].tech[te].load;
                                builds.build[i].floor[fl].fl_req.swch_speed += builds.build[i].floor[fl].dep[de].workers * builds.build[i].floor[fl].dep[de].tech[te].load;
                            }
                            // если отдел использует тыренет
                            if (builds.build[i].floor[fl].dep[de].tech[te].rem_serv)
                            {
                                builds.load[i] += builds.build[i].floor[fl].dep[de].workers * builds.build[i].floor[fl].dep[de].tech[te].load;
                                builds.build[i].floor[fl].fl_req.swch_speed += builds.build[i].floor[fl].dep[de].workers * builds.build[i].floor[fl].dep[de].tech[te].load;
                            }
                        }

                    }
                }
            }
        }
        public static void connect_load_calculation(ref Connect connect, Builds builds)
        {
            connect.connect_load = new int[connect.connect_length.GetLength(0), connect.connect_length.GetLength(1)];
            for (int i = 0; i < connect.connect_load.GetLength(0); i++)
            {
                int fl = 0;
                int de = 0;
                int te = 0;
                while (builds.build[i].floor.Length > fl)
                {
                    while (builds.build[i].floor[fl].dep.Length > de)
                    {
                        while (builds.build[i].floor[fl].dep[de].tech.Length > te)
                        {
                            if (builds.build[i].floor[fl].dep[de].tech[te].user)//если клиент
                            {
                                int z = 0, x = 0, c = 0, v = 0;// для всех серверов
                                while (builds.build.Length > z)
                                {
                                    while (builds.build[z].floor.Length > x)
                                    {
                                        while (builds.build[z].floor[x].dep.Length > c)
                                        {
                                            while (builds.build[z].floor[x].dep[c].tech.Length > v)
                                            {
                                                if (builds.build[i].floor[fl].dep[de].tech[te].name.Equals(builds.build[z].floor[x].dep[c].tech[v].name) && builds.build[z].floor[x].dep[c].tech[v].root)
                                                {
                                                    // записать в путь нагрузку
                                                    if (!load_to_way(
                                                            ref connect,
                                                            connect.connect_length,
                                                            i,
                                                            z,
                                                            builds.build[i].floor[fl].dep[de].tech[te].load * builds.build[i].floor[fl].dep[de].workers
                                                    ))
                                                        throw new Exception("путь не найден");
                                                }
                                                v++;
                                            }
                                            c++;
                                        }
                                        x++;
                                    }
                                    z++;
                                }
                            }
                            if (builds.build[i].floor[fl].dep[de].tech[te].root)
                            {
                                int z = 0, x = 0, c = 0, v = 0;// для всех клиентов
                                while (builds.build.Length > z)
                                {
                                    while (builds.build[z].floor.Length > x)
                                    {
                                        while (builds.build[z].floor[x].dep.Length > c)
                                        {
                                            while (builds.build[z].floor[x].dep[c].tech.Length > v)
                                            {
                                                if (builds.build[i].floor[fl].dep[de].tech[te].name.Equals(builds.build[z].floor[x].dep[c].tech[v].name) && builds.build[z].floor[x].dep[c].tech[v].user)
                                                {
                                                    //записать нагрузку
                                                    if (!load_to_way(
                                                        ref connect,
                                                        connect.connect_length,
                                                        i,
                                                        z,
                                                        builds.build[i].floor[fl].dep[de].tech[te].load * builds.build[z].floor[x].dep[c].workers
                                                    ))
                                                        throw new Exception("путь не найден");
                                                }

                                                v++;
                                            }
                                            c++;
                                        }
                                        x++;
                                    }
                                    z++;
                                }
                            }
                            te++;
                        }
                        de++;
                    }
                    fl++;
                }

            }

        }

        /// <summary>
        /// Функция проставления нагрузки связь между зданиями
        /// </summary>
        /// <param name="connect">объект связи</param>
        /// <param name="connect_length_copy">копия матрицы связей, нужна 
        /// для того чтобы можно было пройтись поиском в глубину пути до 
        /// здания и подсчитать итоговое расстояние</param>
        /// <param name="build_1">номер первого здания</param>
        /// <param name="build_2">номер второго здания</param>
        /// <param name="load">нагрузка</param>
        /// <returns>bool - смог ли проставить нагрузку</returns>        
        public static bool load_to_way(ref Connect connect, int[,] connect_length_copy, int build_1, int build_2, int load) 
        {
            /* 
                Функция просавления нагрузки на кабеля между зданиями.
                Здесь реализован поиск связи между зданиями по графу связей в глубину.
                Подробнее об алгоритме: https://ru.wikipedia.org/wiki/%D0%9F%D0%BE%D0%B8%D1%81%D0%BA_%D0%B2_%D0%B3%D0%BB%D1%83%D0%B1%D0%B8%D0%BD%D1%83

                Рекурсивная фукнция, просматривает все связи пока не дойдет до нужной вершины, если 
                вершина будет найдена в вызывающую функцию вернется true, и будет проставлена нагрузка
                на все ребра пути до здания. Если связь не будет найдена - то тогда в вызывающую функцию вернется false.
            */
            // искомая вершина = вершине старта - мы достигли цели
            if (build_1 == build_2)
                return true;
            // иначе мы начиаем просматривать все оставшиеся связи данной вершины
            for (int i = 0; i < connect_length_copy.GetLength(0); i++)
            {
                // если первое здание с iым соединены
                if (connect_length_copy[build_1, i] != 0)
                {
                    // разъединяем первое здание с iым (проставляем, что эту связь мы посетили)
                    connect_length_copy[build_1, i] = 0;
                    connect_length_copy[i, build_1] = 0;
                    // если iое здание связано с конечным
                    if (load_to_way(ref connect, connect_length_copy, i, build_2, load))
                    {
                        // проставляем нагрузку между зданиями
                        connect.connect_load[build_1, i] += load;
                        connect.connect_load[i, build_1] += load;
                        // цель достигнута мы достигли искомой вершины
                        return true;
                    }
                }
            }
            return false;
        }
        /// <summary>
        /// высчитывание требований этажа
        /// </summary>
        /// <param name="builds">ссылка на здание</param>
        /// <param name="build">номер здания</param>
        /// <param name="floor">номер этажа</param>
        /// <returns></returns>
        public static Floor_req floor_req(ref Builds builds, int build, int floor)
        {
            Floor_req fl = new Floor_req();
            int works = 0, serv=0;
            for (int i = 0; i < builds.build[build].floor[floor].dep.Length; i++)
            {
                works += builds.build[build].floor[floor].dep[i].workers;
                for (int j = 0; j < builds.build[build].floor[floor].dep[i].tech.Length; j++)
                    if (builds.build[build].floor[floor].dep[i].tech[j].root)
                        serv++;

            }
            fl.swch_num = Convert.ToInt32(works / swch_ch_max) + 1;
            fl.swch_chan = Convert.ToInt32(works / fl.swch_num);
            // круговая ссылка wtf???
            // UPD: оказывается это значение уже высчитанов функции build_load_calculation
            fl.swch_speed = builds.build[build].floor[floor].fl_req.swch_speed;
            fl.cabel_length = ((Convert.ToInt32(Math.Sqrt(builds.build[build].square)) * 2 + 1) / 2 + 2) * works / fl.swch_num+ Convert.ToInt32(Math.Sqrt(builds.build[build].square))*2;
            fl.cab_can_length = Convert.ToInt32(fl.cabel_length * 1.2);
            fl.power_uninter_req = fl.swch_num * swch_pow + serv * serv_pow;
            fl.power_req = fl.power_uninter_req + works * pc_pow;
            builds.build[build].floor[floor].fl_req = fl;
            return fl;
        }
        public static Build_req build_req(ref Builds builds, int build)
        {
            Build_req bu = new Build_req();
            for (int i = 1; i <= builds.build[build].floor_num; i++)
                bu.cabel_length += builds.build[build].height * i;
            bu.cab_can_length =Convert.ToInt32(bu.cabel_length * 1.2);
            bu.swch_num = Convert.ToInt32(builds.build[build].floor_num / swch_ch_max) + 1;
            bu.power_uninter_req +=bu.swch_num*swch_pow; ///??????
            bu.swch_chan = Convert.ToInt32(builds.build[build].floor_num / bu.swch_num);
            

            builds.build[build].build_req = bu;
            return bu;
        }


        public static IP[] IP_ret(ref Builds a)
        {
            IP[] ret = new IP[1];
            int cur_str = 0;
            Int64 curr_ip = 0;
            for (int bu = 0; bu < a.build.Length; bu++)
            {
                Console.WriteLine("\nЗдание номер " + (bu + 1));
                Console.Write("Служебные IP " + IPtoSTR(curr_ip));
                ret[cur_str] = new IP();
                ret[cur_str].build_num = bu;
                ret[cur_str].dep_name = "Служебные IP здания" + bu;
                ret[cur_str].adress_start = IPtoSTR(curr_ip);
                curr_ip += a.build[bu].build_req.swch_num + 10;
                Console.WriteLine(" по " + IPtoSTR(curr_ip));
                ret[cur_str].adress_end = IPtoSTR(curr_ip);
                cur_str+=1;
                curr_ip++;
                Array.Resize(ref ret, cur_str+1);
                for (int fl = 0; fl < a.build[bu].floor_num; fl++)
                {
                    Console.WriteLine("   Этаж номер " + (fl + 1));
                    Console.Write("   Служебные IP " + IPtoSTR(curr_ip));
                    ret[cur_str] = new IP();
                    ret[cur_str].build_num = bu;
                    ret[cur_str].dep_name = "   Служебные IP здания " + bu + " этажа " + fl;
                    ret[cur_str].adress_start = IPtoSTR(curr_ip);
                    curr_ip += a.build[bu].floor[fl].fl_req.swch_num + 10;
                    Console.WriteLine(" по " + IPtoSTR(curr_ip));
                    ret[cur_str].adress_end = IPtoSTR(curr_ip);
                    cur_str++;
                    curr_ip++;
                    Array.Resize(ref ret, cur_str+1);
                    for (int depp = 0; depp < a.build[bu].floor[fl].dep_num; depp++)
                    {
                        Console.WriteLine("     Департамент " + a.build[bu].floor[fl].dep[depp].name);
                        Console.Write("      IP    c     " + IPtoSTR(curr_ip));
                        ret[cur_str] = new IP();
                        ret[cur_str].build_num = bu;
                        ret[cur_str].dep_name = a.build[bu].floor[fl].dep[depp].name;
                        ret[cur_str].adress_start = IPtoSTR(curr_ip);
                        curr_ip += a.build[bu].floor[fl].dep[depp].workers;
                        Console.WriteLine(" по " + IPtoSTR(curr_ip));
                        ret[cur_str].adress_end = IPtoSTR(curr_ip);
                        cur_str++;
                        curr_ip++;
                        Array.Resize(ref ret, cur_str+1);
                    }
                }
            }
            return ret;
        }
        public static string IPtoSTR(Int64 ip)
        {
            string outt = "";
            if (ip > 4228250625)
                return null;
            outt += Convert.ToInt64(ip / 256 / 256 / 256);
            outt += ".";
            ip = ip % 256;
            outt += Convert.ToInt64(ip / 256 / 256);
            outt += ".";
            ip = ip % 256;
            outt += Convert.ToInt64(ip / 256);
            outt += ".";
            ip = ip % 256;
            outt += Convert.ToInt64(ip);
            return outt;

        }

        public static void Example()
        { 
            Builds builds = new Builds {
                build = new Build[] {
                    new Build {
                        cabel_type= "keke",
                        build_req = new Build_req { },
                        floor = new Floor[] {
                            new Floor {
                                dep_num = 1,
                                dep = new Department[] {
                                    new Department {
                                        name = "",
                                        workers = 13,
                                        techno_count = 3,
                                        tech = new Technology[] {
                                            new Technology {
                                                load = 3,
                                                rem_serv = true,
                                                // и так далее
                                            }
                                        },
                                    }
                                },
                                Mob_st = 1,
                                fl_req = new Floor_req { }
                            },
                        },
                    }
                }
            };
        }
    }


}
