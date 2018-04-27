using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Amib.Threading;

namespace PdfReporting.Logic
{
    public class ManagedThreadPool : SmartThreadPool
    {
        public ManagedThreadPool(STPStartInfo stpStartInfo) : base(stpStartInfo)
        {
        }

        public static ManagedThreadPool GetStaThreadPool()
        {
            var sTPStartInfo = new STPStartInfo();
            sTPStartInfo.ApartmentState = ApartmentState.STA;
            return new ManagedThreadPool(sTPStartInfo);
        }

        public void FillWith<T>(Action<T> method, IEnumerable<T> parameterList)
        {
            parameterList.ToList().ForEach(parameter => this.QueueWorkItem(method, parameter));
        }
    }
}
