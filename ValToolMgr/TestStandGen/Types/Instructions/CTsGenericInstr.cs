
namespace TestStandGen.Types.Instructions
{
    abstract class CTsGenericInstr
    {
        private static int idSalt = 0;
        private string guid;

        public bool Skipped;
        public bool ForceFailed;
        public bool ForcePassed;

        public abstract string InstructionName { get; protected set; }

        public string GUID 
        { 
            get
            {
                return guid;
            }

            protected set
            {

            }
        }

        protected CTsGenericInstr()
        {
            idSalt++;
            this.guid = idSalt.ToString("0000000000000000000000");
        }

        /// <summary>
        /// Text to display
        /// </summary>
        public string Text;

        public static void resetIdCounter()
        {
            idSalt = 0;
        }
    }
}
