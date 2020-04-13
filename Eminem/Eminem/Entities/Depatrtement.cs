using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    /// <summary>
    /// Депратамент. Содержит в себе ссылки на ифнормационные системы, с которыми работает.
    /// </summary>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Name">Название отдела</param>
        /// <param name="StaffCount">Количество сотрудников</param>
        public Departament(string Name, int StaffCount)
        {
            this.Name = Name;
            this.StaffCount = StaffCount;
            this.InformationSystemList = new List<Tuple<InformationSystem, InformationSystemSettings>>();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Name">Название отдела</param>
        /// <param name="StaffCount">Количество сотрудников</param>
        /// <param name="InformationSystemList">Список используемых информационных систем</param>
        public Departament(string Name, int StaffCount, List<Tuple<InformationSystem, InformationSystemSettings>> InformationSystemList) 
            : this(Name, StaffCount)
        {
            this.AddInformationSystem(InformationSystemList);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Name">Название отдела</param>
        /// <param name="StaffCount">Количество сотрудников</param>
        /// <param name="InformationSystem">Используемая ИС</param>
        /// <param name="isRoot">Является ли отдел корневым для данной ИС</param>
        public Departament(string Name, int StaffCount, Tuple<InformationSystem, InformationSystemSettings> InformationSystem, bool isRoot)
           : this(Name, StaffCount)
        {
            this.AddInformationSystem(InformationSystem.Item1, InformationSystem.Item2, isRoot);
        }
        /// <summary>
        /// Добавить список ИС к департаменту
        /// </summary>
        /// <param name="informationSystem"></param>
        public void AddInformationSystem(List<Tuple<InformationSystem, InformationSystemSettings>> tuples)
        {
            foreach (Tuple<InformationSystem, InformationSystemSettings> tuple in tuples)
                AddInformationSystem(tuple.Item1, tuple.Item2);
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
        /// <summary>
        /// Добавить новую ИС к департаменту
        /// </summary>
        /// <param name="informationSystem"></param>
        public void AddInformationSystem(InformationSystem informationSystem, InformationSystemSettings informationSystemSettings, bool isRoot)
        {
            this.AddInformationSystem(informationSystem, informationSystemSettings);
            if (isRoot)
                informationSystem.Root = this;
        }
        /// <summary>
        /// Количество информационных систем, которыми управляет отдел
        /// </summary>
        /// <returns></returns>
        public int GetRootCount()
        {
            return InformationSystemList.Select(x => x.Item1).Where(x => x.Root == this).Count();
        }
        /// <summary>
        /// Необходимая скорость переключателя для отдела
        /// </summary>
        /// <returns></returns>
        public int GetSwitchSpeed()
        {
            return InformationSystemList.Sum(
                system =>
                {
                    if (system.Item1.Root == this)  // если этот отдел корневой
                        // считаем всех пользователей системы
                        return system.Item1.Traffic * system.Item1.GetClientsCount();
                    else
                        // иначе считаем только сотрудников этого отдела
                        return StaffCount * system.Item1.Traffic;

                }
           );
        }
    }

    

    
}
