using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace OfficeToHtml5Service
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface IService1
    {
        [OperationContract]
        [WebInvoke(Method = "POST",
            BodyStyle = WebMessageBodyStyle.Wrapped,
            ResponseFormat = WebMessageFormat.Json,
            RequestFormat = WebMessageFormat.Json,
            UriTemplate = "getOfficeHtml5View")]
        [return: MessageParameter(Name = "officeConverstionDataObj")]
        officeConverstionDataObj getOfficeHtml5View(officeConverstionDataObj postOfficeObj);
    }


    [DataContract]
    public class officeConverstionDataObj
    {
        [DataMember]
        public string theFileUrl { get; set; }
        [DataMember]
        public string convertedHtml5FileUrl { get; set; }
        [DataMember]
        public string error { get; set; }
        [DataMember]
        public string debugMsg { get; set; }
        [DataMember]
        public string fileContext { get; set; }
    }
}
