using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    public class InformationSystem
    {
        public string Name;
        /// <summary>
        /// нагрузка в Мб
        /// </summary>
        public int Traffic;
        /// <summary>
        /// корневой отдел
        /// </summary>
        public Departament Root;
        /// <summary>
        /// список всех департаментов, использующих эту ИС
        /// </summary>
        public List<Tuple<Departament, InformationSystemSettings>> DepatrtementsList { get; set; }

        public int GetClientsCount()
        {
            return DepatrtementsList.Sum(x => x.Item1.StaffCount);
        }
    }
}
