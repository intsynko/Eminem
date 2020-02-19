using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    public class Floor
    {
        /// <summary>
        /// Список департаментов
        /// </summary>
        public List<Departament> DepartamentsList { get; set; }
        /// <summary>
        /// Количество мобильных станций
        /// </summary>
        public int MobileStationCount;

        // requirements
        public int cabel_length;
        public int cab_can_length;
        public int swch_num;
        public int swch_chan;
        public int swch_speed;
        public int power_req;
        public int power_uninter_req;

        public Floor(int MobileStationCount)
        {
            this.MobileStationCount = MobileStationCount;
        }

        public void CalulateRequrements(int sqare)
        {
            int staffCount = DepartamentsList.Sum(x => x.StaffCount);
            int rootCount = DepartamentsList.Sum(x => x.GetRootCount());
            // defaults
            int swch_pow = 50;
            int serv_pow = 1000;
            int pc_pow = 500;
            int swch_ch_max = 50;

            swch_num = Convert.ToInt32(staffCount / swch_ch_max) + 1;
            swch_chan = Convert.ToInt32(staffCount / swch_num);
            swch_speed = DepartamentsList.Sum(x=>x.GetSwitchSpeed());
            cabel_length = ((Convert.ToInt32(Math.Sqrt(sqare)) * 2 + 1) / 2 + 2) * staffCount / swch_num + Convert.ToInt32(Math.Sqrt(sqare)) * 2;
            cab_can_length = Convert.ToInt32(cabel_length * 1.2);
            power_uninter_req = swch_num * swch_pow + rootCount * serv_pow;
            power_req = power_uninter_req + staffCount * pc_pow;
        }
    }
}
