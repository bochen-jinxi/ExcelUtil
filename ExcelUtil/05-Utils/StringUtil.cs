using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelUtil._05_Utils
{
    public static class StringUtil
    {

        public static int ASCII(this string character)
        {
            if (character.Length != 1)
                throw new Exception("字符无效");
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                int intAsciiCode = (int) asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
        }
    }
}
