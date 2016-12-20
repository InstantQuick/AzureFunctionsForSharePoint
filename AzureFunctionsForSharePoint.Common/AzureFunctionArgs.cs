using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureFunctionsForSharePoint.Common
{
    public abstract class AzureFunctionArgs
    {
        public virtual string StorageAccount { get; set; }
        public virtual string StorageAccountKey { get; set; }
    }
}
