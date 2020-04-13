using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    /// <summary>
    /// Этаж здания
    /// </summary>
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

        /// <summary>
        /// Общая длина кабелей на этаже
        /// </summary>
        public int cabel_length;
        /// <summary>
        /// Общая длина Кабель канала на этаже
        /// </summary>
        public int cab_can_length;
        /// <summary>
        /// Количество свитчей на этаже
        /// </summary>
        public int swch_num;
        /// <summary>
        /// Количество каналов свитчей на этаже
        /// </summary>
        public int swch_chan;
        /// <summary>
        /// Величина скорости свитчей на этаже
        /// </summary>
        public int swch_speed;
        /// <summary>
        /// Величина мощности питания
        /// </summary>
        public int power_req;
        /// <summary>
        /// Величина мощности бесперебойного питания
        /// </summary>
        public int power_uninter_req;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="MobileStationCount">Количество мобильных станций</param>
        public Floor(int MobileStationCount)
        {
            this.MobileStationCount = MobileStationCount;
        }

        /// <summary>
        /// Посчитать требования для этажа
        /// </summary>
        /// <param name="sqare"></param>
        public void CalulateRequrements(int sqare)
        {
            int staffCount = DepartamentsList.Sum(x => x.StaffCount);
            int rootCount = DepartamentsList.Sum(x => x.GetRootCount());

            swch_num = Convert.ToInt32(staffCount / Defaults.swch_ch_max) + 1;
            swch_chan = Convert.ToInt32(staffCount / swch_num);
            swch_speed = DepartamentsList.Sum(x=>x.GetSwitchSpeed());
            cabel_length = (((int)Math.Sqrt(sqare) * 2 + 1) / 2 + 2) * staffCount / swch_num + (int)Math.Sqrt(sqare) * 2;
            cab_can_length = Convert.ToInt32(cabel_length * 1.2);
            power_uninter_req = swch_num * Defaults.swch_pow + rootCount * Defaults.serv_pow;
            power_req = power_uninter_req + staffCount * Defaults.pc_pow;
        }
    }
}
