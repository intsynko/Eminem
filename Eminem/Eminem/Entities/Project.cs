using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eminem.Entities
{
    /// <summary>
    /// Проект. Объединяет в себе все здания.
    /// </summary>
    class Project
    {
        /// <summary>
        /// Список всех, входящих в проект зданий
        /// </summary>
        public List<Build> BuildsList;
        /// <summary>
        /// Матрица смежности связей зданий.
        /// </summary>
        private int[,] _distanseBetweenBuilds;


        // reqirements

        /// <summary>
        /// Нагрузка на проект в целом
        /// </summary>
        public int load;
        /// <summary>
        /// Всего сотрудников
        /// </summary>
        public int all_workers;


        public Project()
        {

        }

        private void ResizeArray<T>(ref T[,] original, int newCoNum, int newRoNum)
        {
            var newArray = new T[newCoNum, newRoNum];
            int columnCount = original.GetLength(1);
            int columnCount2 = newRoNum;
            int columns = original.GetUpperBound(0);
            for (int co = 0; co <= columns; co++)
                Array.Copy(original, co * columnCount, newArray, co * columnCount2, columnCount);
            original = newArray;
        }
        public void AddConnection(Build build1, Build build2, int distanse)
        {
            this.AddConnection(BuildsList.IndexOf(build1), BuildsList.IndexOf(build2), distanse);
        }

        public void AddConnection(int build1, int build2, int distance)
        {
            int size = BuildsList.Count();
            if (_distanseBetweenBuilds == null)
                _distanseBetweenBuilds = new int[size, size];
            if (_distanseBetweenBuilds.GetLength(0) != size)
                ResizeArray<int>(ref _distanseBetweenBuilds, size, size);

            _distanseBetweenBuilds[build1, build2] = distance;
            _distanseBetweenBuilds[build2, build1] = distance;


        }


        public void CalulateRequrements()
        {
            foreach (Build build in BuildsList)
                build.CalulateRequrements();

            load = BuildsList.Sum(build => build.load);
            all_workers = BuildsList.Sum(build => build.FloorList.Sum(floor => floor.DepartamentsList.Sum(dept => dept.StaffCount)));
        }
    }
}
