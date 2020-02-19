using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    public class Departament
    {
        /// <summary>
        ///  Название отдела
        /// </summary>
        public string Name;
        /// <summary>
        ///  Количество сотрудников
        /// </summary>
        public int StaffCount;
        /// <summary>
        /// список информационных систем, используемых в отделе
        /// </summary>
        public List<Tuple<InformationSystem, InformationSystemSettings>> InformationSystemList { get; private set; } 


        public Departament(string Name, int StaffCount)
        {
            this.Name = Name;
            this.StaffCount = StaffCount;
        }
        /// <summary>
        /// Добавить новую ИС к департаменту
        /// </summary>
        /// <param name="informationSystem"></param>
        public void AddInformationSystem(InformationSystem informationSystem, InformationSystemSettings informationSystemSettings)
        {
            InformationSystemList.Add(new Tuple<InformationSystem, InformationSystemSettings>(informationSystem, informationSystemSettings));
            informationSystem.DepatrtementsList.Add(new Tuple<Departament,InformationSystemSettings>(this, informationSystemSettings));
        }
        public int GetRootCount()
        {
            return InformationSystemList.Select(x => x.Item1).Where(x => x.Root == this).Count();
        }
        public int GetSwitchSpeed()
        {
            return InformationSystemList.Sum(
                system => (
                    system.Item1.Root == this ? // если этот отдел корневой
                    system.Item1.Traffic * system.Item1.GetClientsCount() : // считаем всех пользователей системы
                    StaffCount * system.Item1.Traffic // иначе считаем только сотрудников этого отдела
                )
           );
        }
    }

    

    
}
