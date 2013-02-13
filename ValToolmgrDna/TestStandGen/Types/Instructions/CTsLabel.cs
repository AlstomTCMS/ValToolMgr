using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestStandGen.Types.Instructions
{
    //    public enum categoryList
    //    {
    //TS_FORCE,
    //TS_UNFORCE,
    //TS_TEST,
    //TS_WAIT,
    //TS_LABEL,
    //UNKNOWN
    //}

    //        Private Const C_TS_FORCE As String = "CB_Force"
    //Private Const C_TS_UNFORCE As String = "CB_UnForce"
    //Private Const C_TS_TEST As String = "CB_Test"
    //Private Const C_TS_WAIT As String = "NI_Wait"
    //Private Const C_TS_LABEL As String = "Label"
    //Private Const C_UNKNOWN As String = "UNKNOWN"



    /// <summary>
    /// Class that handles Label type.
    /// </summary>
    class CTsLabel : CTsGenericInstr
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="text">Text to be displayed</param>
        public CTsLabel(string text)
        {
            this.Text = text;
        }

        /// <summary>
        /// Static-like name of sequence
        /// </summary>
        public override string InstructionName
        {
            get { return "Label"; }
            protected set { }
        }
    }
}
