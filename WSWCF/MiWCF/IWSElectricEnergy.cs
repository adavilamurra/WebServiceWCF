using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace MiWCF
{
    [ServiceContract]
    public interface IWSElectricEnergy
    {
        [OperationContract]
        EEData GetData(string information);
    }

    [DataContract]
    public class EEData
    {
        string type = "Unknown.";
        string data = "Unknown.";

        public EEData(string info)
        {
            type = info;
        }
        public EEData(string info, string num)
        {
            type = info;
            data = num;
        }

        [DataMember]
        public string Type
        {
            get { return type; }
            set { type = value; }
        }
        [DataMember]
        public string Data
        {
            get { return data; }
            set { data = value; }
        }


    }
}
