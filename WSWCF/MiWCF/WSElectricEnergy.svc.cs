using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace MiWCF
{
    public class WSElectricEnergy : IWSElectricEnergy
    {
        public EEData GetData(string information)
        {
            return new EEData(information);
        }
        public EEData GetData(string information, string number)
        {
            return new EEData(information, number);
        }
    }
}
