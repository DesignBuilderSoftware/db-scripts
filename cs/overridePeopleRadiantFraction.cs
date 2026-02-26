/*
Override radiant heat gain by people.

DesignBuilder always applies default 0.3.

*/

using System.Collections.Generic;
using DB.Extensibility.Contracts;
using EpNet;

namespace DB.Extensibility.Scripts
{
    public class IdfFindAndReplace : ScriptBase, IScript
    {
        public override void BeforeEnergySimulation()
        {
            IdfReader idfReader = new IdfReader(
                ApiEnvironment.EnergyPlusInputIdfPath,
                ApiEnvironment.EnergyPlusInputIddPath);

            const double radiantFraction = 0.15;

            IEnumerable<IdfObject> peopleObjects = idfReader["People"];
            foreach (IdfObject people in peopleObjects)
            {
                people["Fraction Radiant"].Number = radiantFraction;
            }

            idfReader.Save();
        }
    }
}
