using System.Runtime.InteropServices;
using System.Text;

namespace parking_lot_app.Model.MyView
{
    public class IniApi
    {
        public StringBuilder LpReturnedString { get; }
        public string FilePath { get; }
        public int BufferSize { get; }

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string lpString, string lpFileName);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

        public IniApi(string iniPath, int bufferSize = 512)
        {
            FilePath = iniPath;
            BufferSize = bufferSize;
            LpReturnedString = new StringBuilder(bufferSize);
        }

        // read ini date depend on section and key
        public string ReadIniFile(string section, string key, string defaultValue)
        {
            _ = LpReturnedString.Clear();
            _ = GetPrivateProfileString(section, key, defaultValue, LpReturnedString, BufferSize, FilePath);
            return LpReturnedString.ToString();
        }

        // write ini data depend on section and key
        public void WriteIniFile(string section, string key, string value)
        {
            _ = WritePrivateProfileString(section, key, value, FilePath);
        }
    }
}