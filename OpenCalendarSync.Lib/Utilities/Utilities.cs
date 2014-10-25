using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using WallF.BaseNEncodings;

namespace OpenCalendarSync.Lib.Utilities
{
    public static class Defines
    {
        public const string EmailRegularExpression = @"[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?";
    }

    public static class EnumHelper
    {
        /// <summary>
        /// Gets an attribute on an enum field value
        /// </summary>
        /// <typeparam name="T">The type of the attribute you want to retrieve</typeparam>
        /// <param name="enumVal">The enum value</param>
        /// <returns>The attribute of type T that exists on the enum value</returns>
        public static T GetAttributeOfType<T>(this Enum enumVal) where T : Attribute
        {
            var type = enumVal.GetType();
            var memInfo = type.GetMember(enumVal.ToString());
            var attributes = memInfo[0].GetCustomAttributes(typeof(T), false);
            return (attributes.Length > 0) ? (T)attributes[0] : null;
        }
    }

    public static class StringHelper
    {
        public static byte[] GetBytes(string str)
        {
            var bytes = new byte[str.Length * sizeof(char)];
            Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
            return bytes;
        }

        public static string GetString(byte[] bytes)
        {
            var chars = new char[bytes.Length / sizeof(char)];
            Buffer.BlockCopy(bytes, 0, chars, 0, bytes.Length);
            return new string(chars);
        }

        private static readonly char[] Alphabet = "0123456789abcdefghijklmnopqrstuv".ToCharArray();
        private const char Padding = '\0';
        private static readonly BaseEncoding googleBase32 = new Base32Encoding(Alphabet, Padding, "GoogleBase32Enconding");

        public static BaseEncoding GoogleBase32
        {
            get { return googleBase32; }
        }
    }

    public static class MethodHelper
    {
        [MethodImpl(MethodImplOptions.NoInlining)]
        public static string GetCurrentMethod()
        {
            var st = new StackTrace();
            var sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }
    }

    public static class VersionHelper
    {
        public static string LibraryVersion()
        {
            return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        public static DateTime LibraryBuildTime()
        {
            return GetBuildTimeForAssembly(System.Reflection.Assembly.GetExecutingAssembly().Location);
        }

        public static string ExecutingAssemblyVersion()
        {
            return System.Reflection.Assembly.GetCallingAssembly().GetName().Version.ToString();
        }

        public static DateTime ExecutingAssemblyBuildTime()
        {
            return GetBuildTimeForAssembly(System.Reflection.Assembly.GetCallingAssembly().Location);
        }

        private static DateTime GetBuildTimeForAssembly(string assemblyPath)
        {
            const int cPeHeaderOffset = 60;
            const int cLinkerTimestampOffset = 8;
            var b = new byte[2048];
            System.IO.Stream s = null;

            try
            {
                s = new System.IO.FileStream(assemblyPath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                s.Read(b, 0, 2048);
            }
            finally
            {
                if (s != null)
                {
                    s.Close();
                }
            }

            var i = BitConverter.ToInt32(b, cPeHeaderOffset);
            var secondsSince1970 = BitConverter.ToInt32(b, i + cLinkerTimestampOffset);
            var dt = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            dt = dt.AddSeconds(secondsSince1970);
            dt = dt.ToLocalTime();
            return dt;
        }
    }
}