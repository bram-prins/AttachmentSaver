using System;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttachmentSaver
{
    public class OpslagProfiel
    {
        public string Name { get; set; }
        public string Sender { get; set; }
        public string FilenameFilter { get; set; }
        public List<string> FileExtensions { get; set; }
        public string FileLocation { get; set; }
    }
}
