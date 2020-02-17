using System;
using System.Collections.Generic;
using System.Text;



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
    public struct Build //описание здания
    {
        public int floor_num;
        public int square;
        public int hight;
        public Floor[] floor; //создаем массив с кол-вом этажей и их перечисленим в каждой ячейке
        public Build_req build_req;

    }
    public struct Floor
    {
        public Departement[] dep; //создаем массив с кол-вом отделов на этаже и их перечисленим в каждой ячейке
        public int Mob_st;
        public Floor_req fl_req;
    }
    public struct Departement
    {
        public String name;
        public int workers;
        public Techology[] tech; //список технологий, используемых в отделе
    }
    public struct Techology
    {
        public string name;
        public int load; // нагрузка в Мб!!!!
        public bool root; //является ли отдел корневым отделом, к которому идут все запросы
        public bool user; //является ли отдел пользователем технологии
        public bool rem_serv; //использует ли программа ресурсы сети интернет, если да - все запросы считаются внесетевыми и не расчитываются между зданиями
    }
    public struct Floor_req //только для расчета на этаже
    {
        public int cabel_length;
        public int cab_can_length;
        public int swch_num;
        public int swch_chan;
        public int swch_speed;
        public int power_req;
        public int power_uninter_req;
    }
    public struct Build_req //только для расчета между этажами
    {
        public int cabel_length;
        public int cab_can_length;
        public int swch_num; //возможно расположение некоторых свичей на разных этажах, при 2+ свичах третий с магистрального допом
        public int swch_chan;
        public int power_uninter_req; //только для главных серверных
    }
    public struct Connect
    {
        public int[,] connect_length;//дб КВАДРАТНЫМ!!!
        public void add_connect(int first_build, int second_build, int length)
        {
            if (connect_length.GetLength(0)> second_build && connect_length.GetLength(0) > first_build)
            {
                connect_length[first_build,second_build] = length;
                connect_length[second_build,first_build] = length;
            }
            else throw new Exception("Такого здания не существует");
        }
        public int[,] connect_load;
   
    }
    public class Method
    {
        static int swch_pow = 50;
        static int serv_pow = 1000;
        static int pc_pow = 500;
        static int swch_ch_max = 50;






        static void build_load_calculation(ref Builds builds)
        {
            int i = 0;
            builds.load = new int[builds.build.Length];//создаем массив с нагрузками на здания
            while (builds.build.Length > i)
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
                            if (builds.build[i].floor[fl].dep[de].tech[te].root)
                            {
                                int z = 0, x = 0, c = 0, v = 0, client = 0;
                                while (builds.build.Length > z)//считаем кол-во клиентов технологии
                                {
                                    while (builds.build[z].floor.Length > x)
                                    {
                                        while (builds.build[z].floor[x].dep.Length > c)
                                        {
                                            while (builds.build[z].floor[x].dep[c].tech.Length > v)
                                            {
                                                if (builds.build[i].floor[fl].dep[de].tech[te].name.Equals(builds.build[z].floor[x].dep[c].tech[v].name)&& builds.build[z].floor[x].dep[c].tech[v].user)
                                                    client += builds.build[z].floor[x].dep[c].workers;//подсчет клиентов технологии
                                                v++;

                                            }
                                            c++;
                                        }
                                        x++;
                                    }
                                    z++;
                                }
                                builds.load[i] += client * builds.build[i].floor[fl].dep[de].tech[te].load;//считаю нагрузку на сервера
                                builds.build[i].floor[fl].fl_req.swch_speed+= client * builds.build[i].floor[fl].dep[de].tech[te].load;
                            }
                            if(builds.build[i].floor[fl].dep[de].tech[te].user)
                            {
                                builds.load[i] += builds.build[i].floor[fl].dep[de].workers * builds.build[i].floor[fl].dep[de].tech[te].load;//считаю нагрузку клиентов
                                builds.build[i].floor[fl].fl_req.swch_speed += builds.build[i].floor[fl].dep[de].workers * builds.build[i].floor[fl].dep[de].tech[te].load;
                            }
                            if (builds.build[i].floor[fl].dep[de].tech[te].rem_serv)
                            {
                                builds.load[i] += builds.build[i].floor[fl].dep[de].workers * builds.build[i].floor[fl].dep[de].tech[te].load;
                                builds.build[i].floor[fl].fl_req.swch_speed += builds.build[i].floor[fl].dep[de].workers * builds.build[i].floor[fl].dep[de].tech[te].load;
                            }
                            te++;
                        }
                        de++;

                    }fl++;

                }i++;
                
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
                                                        if (builds.build[i].floor[fl].dep[de].tech[te].name.Equals(builds.build[z].floor[x].dep[c].tech[v].name)&& builds.build[z].floor[x].dep[c].tech[v].root)
                                                        {
                                                            if(!load_to_way(ref connect, connect.connect_length, i, z, builds.build[i].floor[fl].dep[de].tech[te].load * builds.build[i].floor[fl].dep[de].workers))//записать в путь нагрузку
                                                                throw new Exception("путь не найден");
                                                        }
                                                    }v++;
                                                }c++;
                                            }x++;
                                        }z++;
                                    }if (builds.build[i].floor[fl].dep[de].tech[te].root)
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
                                                    if (!load_to_way(ref connect, connect.connect_length, i, z, builds.build[i].floor[fl].dep[de].tech[te].load * builds.build[z].floor[x].dep[c].workers))//записать в путь нагрузку
                                                        throw new Exception("путь не найден");
                                                }
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
                    }de++;
                }fl++;

            }

        }
        public static bool load_to_way(ref Connect connect, int[,] co, int a, int b, int load) //метод заполнения пути 
        {
            List<Int32> way = new List<int>();
            bool check = false;
            way.Add(a);
            if (a == b)
                return true;
            for (int i = 0; i < connect.connect_length.GetLength(0); i++)
                if (co[a, i] != 0)
                {
                    co[a, i] = 0;
                    if (load_to_way(ref connect, co, i, b, load))
                    {
                        connect.connect_load[a, i] += load;
                        connect.connect_load[i,a] += load;
                        check = true;
                    }

                }
            return check;
        }
        public static Floor_req floor_req(ref Builds builds, int build, int floor) //для каждого этажа сначала выполниить это!!!!!
        {
            Floor_req fl = new Floor_req(); int works = 0, serv=0;
            for (int i = 0; i < builds.build[build].floor[floor].dep.Length; i++)
            {
                works += builds.build[build].floor[floor].dep[i].workers;
                for (int j = 0; j < builds.build[build].floor[floor].dep[i].tech.Length; j++)
                    if (builds.build[build].floor[floor].dep[i].tech[j].root)
                        serv++;

            }
            fl.swch_num = Convert.ToInt32(works / swch_ch_max) + 1;
            fl.swch_chan = Convert.ToInt32(works / fl.swch_num);
            fl.swch_speed = builds.build[build].floor[floor].fl_req.swch_speed;
            fl.cabel_length = ((Convert.ToInt32(Math.Sqrt(builds.build[build].square)) * 2 + 1) / 2 + 2) * works / fl.swch_num;
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
                bu.cabel_length += builds.build[build].hight * i;
            bu.cab_can_length =Convert.ToInt32(bu.cabel_length * 1.2);
            bu.power_uninter_req = 0; ///??????
            bu.swch_num = Convert.ToInt32(builds.build[build].floor_num / swch_ch_max) + 1;
            bu.swch_chan = Convert.ToInt32(builds.build[build].floor_num / bu.swch_num);
            

            builds.build[build].build_req = bu;
            return bu;
        }



    }


}
