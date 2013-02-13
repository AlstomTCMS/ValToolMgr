
namespace TestStandGen
{
    abstract class CTsGenericInstr
    {

        public abstract string InstrTsName { get; protected set; }

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
