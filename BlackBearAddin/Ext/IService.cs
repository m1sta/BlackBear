using System.ServiceModel;
using System.ServiceModel.Web;

namespace ExcelPythonAddIn.Ext
{
    [ServiceContract]
    public interface IService
    {
        [OperationContract]
        [WebGet]
        string EchoWithGet(string s);

        [OperationContract]
        [WebInvoke]
        string EchoWithPost(string s);
    }

    public class MonacoService : IService
    {
        public string EchoWithGet(string s)
        {
            return "Echo:" + s;
        }

        public string EchoWithPost(string s)
        {
            return "Echo:" + s;
        }

    }
}