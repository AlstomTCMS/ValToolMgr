
namespace TestStandGen.Types.Instructions
{
    abstract class CTsGenericInstr
    {

        public abstract string InstructionName { get; protected set; }

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
