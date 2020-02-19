using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    class Build
    {
        public string CabelType;
        public int Square;
        public int Height;
        public List<Floor> FloorList;

        // requirements
        public int load;
        public int cabel_length;
        public int cab_can_length;
        public int swch_num; // возможно расположение некоторых свичей на разных этажах, при 2+ свичах третий с магистрального допом
        public int swch_chan;
        public int power_uninter_req; // только для главных серверных

        public Build(int Square, int Height)
        {
            this.Square = Square;
            this.Height = Height;
        }

        public void CalulateRequrements()
        {
            // сначала вычисляем требования для каждого этажа
            foreach (Floor floor in FloorList)
                floor.CalulateRequrements(Square);

            // потом вычисляем свои требования
            load = FloorList.Sum(x=>x.swch_speed);
            cabel_length = FloorList.Count() * Height;
            cab_can_length = (int)(cabel_length * 1.2);
            swch_num = FloorList.Count() * Defaults.swch_ch_max;
            power_uninter_req = swch_num * Defaults.swch_pow;
            swch_chan = FloorList.Count() / swch_num;
        }
    }
}
