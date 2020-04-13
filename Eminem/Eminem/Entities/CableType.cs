using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    /// <summary>
    /// Кабель для соединения
    /// </summary>
    class CableType
    {
        public static CableType[] CommonTypes = new CableType[] {
            new CableType("Ethernet", 1000, 100),
            new CableType("Оптоволоконный", 40000, 3000),
            new CableType("Коаксиальный", 10, 500)
        };
        /// <summary>
        /// Название
        /// </summary>
        public string Name;
        /// <summary>
        /// Максимальная нагрузка
        /// </summary>
        public int MaxLoad;
        /// <summary>
        /// Максимальная длина
        /// </summary>
        public int MaxLength;

        /// <summary>
        /// Создать новый тип кабеля
        /// </summary>
        /// <param name="Name">Название</param>
        /// <param name="MaxLoad">Максимальная нагрузка</param>
        /// <param name="MaxLength">Максимальная длина</param>
        public CableType(string Name, int MaxLoad, int MaxLength)
        {
            this.Name = Name;
            this.MaxLoad = MaxLoad;
            this.MaxLength = MaxLength;
        }
    }
}
