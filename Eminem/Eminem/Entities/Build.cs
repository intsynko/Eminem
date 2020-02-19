using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    class Build
    {
        public string Cabel_type;
        public int Floor_num;
        public int Square;
        public int Height;
        public List<Floor> FloorList;

        // requirements
        public int cabel_length;
        public int cab_can_length;
        public int swch_num; //возможно расположение некоторых свичей на разных этажах, при 2+ свичах третий с магистрального допом
        public int swch_chan;
        public int power_uninter_req; //только для главных серверных
    }
}
