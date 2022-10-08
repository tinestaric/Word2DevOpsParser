namespace Word2DevOpsParser
{
    using System.Collections.Generic;

    internal class BacklogItem
    {
        public string StyleName { get; set; }

        public string Name { get; set; }

        public string Content { get; set; }

        public Dictionary<string, string> Pictures { get; internal set; }

        public int Indent { get; internal set; }
    }
}
