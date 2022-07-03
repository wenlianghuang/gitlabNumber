using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SomeDatainLibrary
{
    public enum LogLevel
    {
        True = 0,
        Debug = 1,
        Information = 2,
        Warning = 3,
        Error = 4,
        Fatal = 5,
        FatalWarning = 6,
        FatalError = 7,

    };
    
    public static class ClasswithExcel
    {
        public static string[] ExcelNameofColumn =
        {
            "B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
            "AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ",
            "BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL,","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BA"
        };


    }
}
