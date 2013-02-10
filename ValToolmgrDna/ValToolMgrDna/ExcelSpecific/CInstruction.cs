using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValToolMgrDna.ExcelSpecific
{
    class CInstruction
    {
        public enum actionList
        {
            A_INIT_TASK,
            A_POPUP,
            A_LABEL,
            A_FORCE,
            A_FORCE_NN,
            A_UNFORCE,
            A_UNFORCE_NN,
            A_WRITE,
            A_WRITE_NN,
            A_READ,
            A_READ_NN,
            A_TEST,
            A_TEST_NN,
            A_TEST_ANA,
            A_TEST_ANA_NN,
            A_CALL,
            A_WAIT,
            A_UNFORCE_ARRAY,
            A_UNFORCE_ARRAY_NN,
            A_FORCE_ARRAY_ALL,
            A_FORCE_ARRAY_ALL_NN,
            A_FORCE_ARRAY_ELT,
            A_FORCE_ARRAY_ELT_NN,
            A_QA_RESET_ALL,
            A_QA_UNFORCE_ALL,
            A_QA_FORCE_VAR,
            A_QA_UNFORCE_VAR,
            A_TEST_MMI_150,
            A_TEST_CMX_EVOL,
            A_HMI_START_TEST,
            A_HMI_STOP_TEST,
            A_HMI_SEND_KEY,
            A_STATEMENT,
            UNIMPLEMENTED
        }

        public object data { get; set; }
        public actionList category { get; set; }
    }
}
