using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace Mpcjob.Otp
{
    [Guid("4F5353CE-3647-4C78-B549-3FDF61E8D2E7")]
    [ComVisible(true)]
    public interface IOtpVerifier
    {
        string GenerateSecret();
        bool VerifyOtp(string secret, string code);
    }

    [Guid("D8E53A5A-4DF1-4F43-85B9-F8436B13D5E4")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class OtpVerifier : IOtpVerifier
    {
        // 1. Generate a 16-character Base32 Secret Key
        public string GenerateSecret()
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
            byte[] bytes = new byte[16];
            using (var rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(bytes);
            }
            char[] result = new char[16];
            for (int i = 0; i < 16; i++)
            {
                result[i] = chars[bytes[i] % chars.Length];
            }
            return new string(result);
        }

        // 2. Verify OTP code using TOTP standard algorithm
        public bool VerifyOtp(string secret, string code)
        {
            if (string.IsNullOrEmpty(secret) || string.IsNullOrEmpty(code) || code.Length != 6)
                return false;

            try
            {
                byte[] key = Base32Decode(secret);
                long currentStep = GetCurrentTimeStep();

                // Accept network delay/time offset (current-1, current, current+1)
                for (int i = -1; i <= 1; i++)
                {
                    if (CalculateTotp(key, currentStep + i) == code)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private long GetCurrentTimeStep()
        {
            DateTime epochStart = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            TimeSpan elapsed = DateTime.UtcNow - epochStart;
            return (long)(elapsed.TotalSeconds / 30);
        }

        private string CalculateTotp(byte[] key, long step)
        {
            byte[] stepBytes = BitConverter.GetBytes(step);
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(stepBytes);
            }

            using (var hmac = new HMACSHA1(key))
            {
                byte[] hash = hmac.ComputeHash(stepBytes);
                int offset = hash[hash.Length - 1] & 0x0F;
                int binary = ((hash[offset] & 0x7F) << 24)
                           | ((hash[offset + 1] & 0xFF) << 16)
                           | ((hash[offset + 2] & 0xFF) << 8)
                           | (hash[offset + 3] & 0xFF);
                int otp = binary % 1000000;
                return otp.ToString("D6");
            }
        }

        private byte[] Base32Decode(string base32)
        {
            base32 = base32.Trim().ToUpper().Replace("=", "");
            int byteCount = base32.Length * 5 / 8;
            byte[] returnArray = new byte[byteCount];
            byte curByte = 0, bitsRemaining = 8;
            int mask = 0, arrayIndex = 0;
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";

            foreach (char c in base32)
            {
                int value = alphabet.IndexOf(c);
                if (value < 0) continue;

                if (bitsRemaining >= 5)
                {
                    mask = value << (bitsRemaining - 5);
                    curByte = (byte)(curByte | mask);
                    bitsRemaining -= 5;
                }
                else
                {
                    mask = value >> (5 - bitsRemaining);
                    curByte = (byte)(curByte | mask);
                    returnArray[arrayIndex++] = curByte;
                    curByte = (byte)((value << (3 + bitsRemaining)) & 0xFF);
                    bitsRemaining += 3;
                }
            }
            return returnArray;
        }
    }
}
