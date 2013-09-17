using System.Collections.Generic;

namespace ValToolMgrDna.Report
{
    public abstract class ItemReport
    {
        public string Name = "";
        public List<ItemReport> List = new List<ItemReport>();
        public List<MessageReport> Messages = new List<MessageReport>();

        public ItemReport(string name)
        {
            Name = name;
        }

        public void add(ItemReport item)
        {
            List.Add(item);
        }

        public void add(MessageReport message)
        {
            Messages.Add(message);
        }

        public int NbrMessages
        {
            get
            {
                int count = Messages.Count;

                foreach (ItemReport item in List)
                {
                    count += item.NbrMessages;
                }

                return count;
            }
        }
    }
}
