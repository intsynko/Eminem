using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Eminem.Entities;

namespace Eminem
{
    class Test
    {
        public static void TestFunc()
        {
            InformationSystem informationSystem = new InformationSystem("Main", 3);

            Project project = new Project()
            {
                BuildsList = new List<Build>()
                {
                    new Build(300, 3) {
                        FloorList = new List<Floor> {
                            new Floor(1) {
                                DepartamentsList = new List<Departament> {
                                    new Departament(
                                        "Departament 1",
                                        13,
                                        new Tuple<InformationSystem, InformationSystemSettings>(
                                            informationSystem,
                                            new InformationSystemSettings(false)
                                        ),
                                        true
                                    ),
                                    new Departament(
                                        "Departament 2",
                                        5,
                                        new Tuple<InformationSystem, InformationSystemSettings>(
                                            informationSystem,
                                            new InformationSystemSettings(false)
                                        ),
                                        false
                                    ),
                                }
                            }
                        }
                    }
                }
            };
            project.CalulateRequrements();
            Console.ReadLine();
        }
    }
}
