using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    /// <summary>
    /// Тип кабеля
    /// </summary>
    //enum CableType { Ethernet, OpticalFiber, Coaxial }

    /// <summary>
    /// Здание.
    /// </summary>
    class Build
    {
        public CableType CabelType;
        public int Square;
        public int FloorHeight;
        public List<Floor> FloorList;

        // requirements

        /// <summary>
        /// Предполагаемая нагрузка внутри здания
        /// </summary>
        public int load;
        /// <summary>
        /// Итоговая длина кабелей в здании
        /// </summary>
        public int cabel_length;
        /// <summary>
        /// Итоговая длина кабель-канала в здании
        /// </summary>
        public int cab_can_length;
        /// <summary>
        /// Итоговое количество свитчей в здании
        /// </summary>
        public int swch_num; // возможно расположение некоторых свичей на разных этажах, при 2+ свичах третий с магистрального допом
        /// <summary>
        /// Итоговое количество каналов свитчей в здании
        /// </summary>
        public int swch_chan;
        /// <summary>
        /// Итоговая размерность требуемой мощности бесперебойного питания в здании
        /// </summary>
        public int power_uninter_req; // только для главных серверных

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Square">Площадь этажа здания</param>
        /// <param name="FloorHeight">Высота этажа</param>
        public Build(int Square, int FloorHeight)
        {
            this.Square = Square;
            this.FloorHeight = FloorHeight;
        }

        /// <summary>
        /// Посчитать требования для здания
        /// </summary>
        public void CalulateRequrements()
        {
            // сначала вычисляем требования для каждого этажа
            foreach (Floor floor in FloorList)
                floor.CalulateRequrements(Square);

            // потом вычисляем свои требования
            load = FloorList.Sum(x=>x.swch_speed);
            cabel_length = FloorList.Count() * FloorHeight;
            cab_can_length = (int)(cabel_length * 1.2);
            swch_num = FloorList.Count() * Defaults.swch_ch_max;
            power_uninter_req = swch_num * Defaults.swch_pow;
            swch_chan = FloorList.Count() / swch_num;
        }
    }
}
