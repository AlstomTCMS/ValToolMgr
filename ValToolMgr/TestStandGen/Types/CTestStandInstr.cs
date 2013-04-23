
namespace TestStandGen
{
    class CTestStandInstr
    {
        public enum categoryList
        {
    TS_FORCE,
    TS_UNFORCE,
    TS_TEST,
    TS_CALL,
    TS_WAIT,
    TS_LABEL,
    UNKNOWN
    }


        public categoryList category { get; set; }

        public object Data { get; set; }
    }
}
