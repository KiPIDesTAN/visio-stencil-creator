using System.Collections.Generic;

namespace VisioStencilCreator
{
    public class VisioImage
    {
        public int TargetArea { get; set; }
    }

    public class VisioConnection
    {
        public string Name { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
    }

    public class VisioStencilConfig
    {
        public VisioImage Image { get; set; } = null;
        public IList<VisioConnection> Connections { get; set; } = new List<VisioConnection>();
    }

}
