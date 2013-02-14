
namespace TestStandGen.Types.Instructions
{
    abstract class CTsGenericInstr
    {
        private static int idSalt = 0;

        public abstract string InstructionName { get; protected set; }

        public string GUID 
        { 
            get
            {
                idSalt++;
                return idSalt.ToString("0000000000000000000000");
            }

            protected set
            {

            }
        }

        /// <summary>
        /// Parametrize instruction as skipped if true.
        /// </summary>
        public bool SkipInstruction;

        /// <summary>
        /// Text to display
        /// </summary>
        public string Text;
    }
}
