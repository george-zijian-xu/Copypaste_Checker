// TraceEncryption.cs
// AES-256-GCM payload encryption + RSA-OAEP key wrap + HMAC-SHA256 tick chain.
// Uses Windows BCrypt (CNG) via P/Invoke for AES-GCM — no external dependencies.
//
// Encrypted XML format:
//   <pasteTrace xmlns="urn:paste-monitor">
//     <header kv="1" ek="BASE64_RSA_WRAPPED_AES_KEY"/>
//     <payload iv="BASE64_IV" ct="BASE64_CIPHERTEXT" tag="BASE64_TAG"/>
//   </pasteTrace>
//
// The plaintext that goes into AES-GCM is the inner XML (sessions + ticks + btree + pastes).
// Each tick row carries an hmac="" attribute computed as:
//   HMAC[n] = HMAC-SHA256(aesKey, HMAC[n-1] || tickId || op || loc || len || paste)
// The chain root is seeded with: HMAC-SHA256(aesKey, docGuid || sessionId || sessionStartUtc)
//
// Key version "1" = RSA-2048-OAEP-SHA256 wrapping AES-256-GCM.
// The RSA public key is baked in as a DER base64 constant below.
// The server holds the matching private key to decrypt.

using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace IsItYoursWordAddIn
{
    internal static class TraceEncryption
    {
        public const string KeyVersion = "1";

        // RSA-2048 public key in .NET XML format — dev key.
        // Replace with production key before shipping.
        // Matching private key: dev_private.pem (never commit to source control in production).
        // Modulus and Exponent are base64-encoded big-endian integers (no leading 0x00).
        private const string PublicKeyXml =
            "<RSAKeyValue>" +
            "<Modulus>nvJHz/Pm6+cwpNVUtg7lf0jZN6/y6u9+RYPd3SMn8g6SWonewHBC5THJo8jABp1F2SmgNvRUM42P" +
            "qG0UVEPfmNl0Pq7C8o5os8SwO3WliDQQEzXHKKCD8mG0i2PzF3PXRJOSN54KbCd7A6i6OMcawZ" +
            "mppGG5dp+rnVgJiR04QhrKz3QntcDZEiPMxiSHNQ1XeQ5rHWmAk760IRUXNwQRh4bsWw6dsLqC" +
            "xsivVfPitNRNXUVwhHFoAqgen+HPrYON+hM23q2nCSPvoL4fV/ZruV0tnhUO8ZOoWNGBqkbgzD" +
            "ZnF71bc+QpKSIyu1WdpDob0ySu4OQIp1INur2FgcwZiw==</Modulus>" +
            "<Exponent>AQAB</Exponent>" +
            "</RSAKeyValue>";

        // ── Key management ──────────────────────────────────────────────────────

        public static byte[] GenerateAesKey()
        {
            var key = new byte[32];
            using (var rng = new RNGCryptoServiceProvider())
                rng.GetBytes(key);
            return key;
        }

        /// <summary>
        /// Wrap the AES key with the server's RSA public key (OAEP-SHA1, CAPI compatible).
        /// Stored as "ek" in the XML header -- only the server can unwrap this.
        /// </summary>
        public static string WrapAesKey(byte[] aesKey)
        {
            using (var rsa = new RSACryptoServiceProvider())
            {
                rsa.FromXmlString(PublicKeyXml);
                var wrapped = rsa.Encrypt(aesKey, fOAEP: true);
                return Convert.ToBase64String(wrapped);
            }
        }

        /// <summary>
        /// Wrap the AES key with DPAPI (Windows Data Protection API, CurrentUser scope).
        /// Stored as "lk" in the XML header -- only the same Windows user on the same machine can unwrap.
        /// This lets the add-in re-hydrate the trace on document re-open without the server.
        /// entropy: docGuid bytes bind the key to this specific document.
        /// </summary>
        public static string WrapAesKeyLocal(byte[] aesKey, string docGuid)
        {
            var entropy = Encoding.UTF8.GetBytes(docGuid ?? "");
            var wrapped = System.Security.Cryptography.ProtectedData.Protect(
                aesKey, entropy, System.Security.Cryptography.DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(wrapped);
        }

        /// <summary>Unwrap the DPAPI-wrapped AES key. Returns null on failure.</summary>
        public static byte[] UnwrapAesKeyLocal(string lkB64, string docGuid)
        {
            try
            {
                var wrapped = Convert.FromBase64String(lkB64);
                var entropy = Encoding.UTF8.GetBytes(docGuid ?? "");
                return System.Security.Cryptography.ProtectedData.Unprotect(
                    wrapped, entropy, System.Security.Cryptography.DataProtectionScope.CurrentUser);
            }
            catch { return null; }
        }

        // ── AES-256-GCM via Windows BCrypt (CNG) ────────────────────────────────

        public static (string ivB64, string ctB64, string tagB64) Encrypt(byte[] aesKey, string plaintext)
        {
            var iv = new byte[12];
            using (var rng = new RNGCryptoServiceProvider()) rng.GetBytes(iv);

            var pt = Encoding.UTF8.GetBytes(plaintext);
            AesGcmBCrypt.Encrypt(aesKey, iv, pt, out var ct, out var tag);

            return (Convert.ToBase64String(iv), Convert.ToBase64String(ct), Convert.ToBase64String(tag));
        }

        public static string Decrypt(byte[] aesKey, string ivB64, string ctB64, string tagB64)
        {
            var iv = Convert.FromBase64String(ivB64);
            var ct = Convert.FromBase64String(ctB64);
            var tag = Convert.FromBase64String(tagB64);

            var pt = AesGcmBCrypt.Decrypt(aesKey, iv, ct, tag);
            return Encoding.UTF8.GetString(pt);
        }

        // ── HMAC chain ──────────────────────────────────────────────────────────

        public static byte[] ChainRoot(byte[] aesKey, string docGuid, int sessionId, DateTime sessionStartUtc)
        {
            var seed = Encoding.UTF8.GetBytes(
                $"{docGuid}|{sessionId:000}|{sessionStartUtc:yyyy-MM-ddTHH:mm:ssZ}");
            using (var hmac = new HMACSHA256(aesKey))
                return hmac.ComputeHash(seed);
        }

        public static byte[] ChainStep(byte[] aesKey, byte[] prevHmac, string tickId, string op, int loc, int len, int paste)
        {
            var data = Encoding.UTF8.GetBytes(
                $"{Convert.ToBase64String(prevHmac)}|{tickId}|{op}|{loc}|{len}|{paste}");
            using (var hmac = new HMACSHA256(aesKey))
                return hmac.ComputeHash(data);
        }

        public static string HmacToString(byte[] hmac) => Convert.ToBase64String(hmac);
    }

    // ── BCrypt AES-GCM P/Invoke ─────────────────────────────────────────────────
    // Uses Windows CNG (Vista+). No external dependencies.

    internal static class AesGcmBCrypt
    {
        private const string BCRYPT_AES_ALGORITHM = "AES";
        private const string BCRYPT_CHAINING_MODE = "ChainingMode";
        private const string BCRYPT_CHAIN_MODE_GCM = "ChainingModeGCM";
        private const uint BCRYPT_OPEN_ALGORITHM_PROVIDER_FLAGS = 0;
        private const uint STATUS_SUCCESS = 0;

        [DllImport("bcrypt.dll", CharSet = CharSet.Unicode)]
        private static extern uint BCryptOpenAlgorithmProvider(out IntPtr phAlgorithm, string pszAlgId, string pszImplementation, uint dwFlags);

        [DllImport("bcrypt.dll", CharSet = CharSet.Unicode)]
        private static extern uint BCryptSetProperty(IntPtr hObject, string pszProperty, byte[] pbInput, int cbInput, uint dwFlags);

        [DllImport("bcrypt.dll")]
        private static extern uint BCryptGenerateSymmetricKey(IntPtr hAlgorithm, out IntPtr phKey, IntPtr pbKeyObject, int cbKeyObject, byte[] pbSecret, int cbSecret, uint dwFlags);

        [DllImport("bcrypt.dll")]
        private static extern uint BCryptEncrypt(IntPtr hKey, byte[] pbInput, int cbInput, ref BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO pPaddingInfo, byte[] pbIV, int cbIV, byte[] pbOutput, int cbOutput, out int pcbResult, uint dwFlags);

        [DllImport("bcrypt.dll")]
        private static extern uint BCryptDecrypt(IntPtr hKey, byte[] pbInput, int cbInput, ref BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO pPaddingInfo, byte[] pbIV, int cbIV, byte[] pbOutput, int cbOutput, out int pcbResult, uint dwFlags);

        [DllImport("bcrypt.dll")]
        private static extern uint BCryptDestroyKey(IntPtr hKey);

        [DllImport("bcrypt.dll")]
        private static extern uint BCryptCloseAlgorithmProvider(IntPtr hAlgorithm, uint dwFlags);

        [StructLayout(LayoutKind.Sequential)]
        private struct BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
        {
            public uint cbSize;
            public uint dwInfoVersion;
            public IntPtr pbNonce;
            public uint cbNonce;
            public IntPtr pbAuthData;
            public uint cbAuthData;
            public IntPtr pbTag;
            public uint cbTag;
            public IntPtr pbMacContext;
            public uint cbMacContext;
            public uint cbAAD;
            public ulong cbData;
            public uint dwFlags;
        }

        private const uint BCRYPT_AUTH_MODE_CHAIN_CALLS_FLAG = 0x00000001;
        private const uint BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO_VERSION = 1;

        public static void Encrypt(byte[] key, byte[] nonce, byte[] plaintext, out byte[] ciphertext, out byte[] tag)
        {
            tag = new byte[16];
            ciphertext = new byte[plaintext.Length];

            IntPtr hAlg = IntPtr.Zero, hKey = IntPtr.Zero;
            GCHandle nonceHandle = default, tagHandle = default;
            try
            {
                Check(BCryptOpenAlgorithmProvider(out hAlg, BCRYPT_AES_ALGORITHM, null, BCRYPT_OPEN_ALGORITHM_PROVIDER_FLAGS));
                var modeBytes = Encoding.Unicode.GetBytes(BCRYPT_CHAIN_MODE_GCM + "\0");
                Check(BCryptSetProperty(hAlg, BCRYPT_CHAINING_MODE, modeBytes, modeBytes.Length, 0));
                Check(BCryptGenerateSymmetricKey(hAlg, out hKey, IntPtr.Zero, 0, key, key.Length, 0));

                nonceHandle = GCHandle.Alloc(nonce, GCHandleType.Pinned);
                tagHandle = GCHandle.Alloc(tag, GCHandleType.Pinned);

                var info = new BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
                {
                    cbSize = (uint)Marshal.SizeOf(typeof(BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO)),
                    dwInfoVersion = BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO_VERSION,
                    pbNonce = nonceHandle.AddrOfPinnedObject(),
                    cbNonce = (uint)nonce.Length,
                    pbTag = tagHandle.AddrOfPinnedObject(),
                    cbTag = (uint)tag.Length
                };

                Check(BCryptEncrypt(hKey, plaintext, plaintext.Length, ref info, null, 0, ciphertext, ciphertext.Length, out _, 0));
            }
            finally
            {
                if (nonceHandle.IsAllocated) nonceHandle.Free();
                if (tagHandle.IsAllocated) tagHandle.Free();
                if (hKey != IntPtr.Zero) BCryptDestroyKey(hKey);
                if (hAlg != IntPtr.Zero) BCryptCloseAlgorithmProvider(hAlg, 0);
            }
        }

        public static byte[] Decrypt(byte[] key, byte[] nonce, byte[] ciphertext, byte[] tag)
        {
            var plaintext = new byte[ciphertext.Length];
            IntPtr hAlg = IntPtr.Zero, hKey = IntPtr.Zero;
            GCHandle nonceHandle = default, tagHandle = default;
            try
            {
                Check(BCryptOpenAlgorithmProvider(out hAlg, BCRYPT_AES_ALGORITHM, null, BCRYPT_OPEN_ALGORITHM_PROVIDER_FLAGS));
                var modeBytes = Encoding.Unicode.GetBytes(BCRYPT_CHAIN_MODE_GCM + "\0");
                Check(BCryptSetProperty(hAlg, BCRYPT_CHAINING_MODE, modeBytes, modeBytes.Length, 0));
                Check(BCryptGenerateSymmetricKey(hAlg, out hKey, IntPtr.Zero, 0, key, key.Length, 0));

                nonceHandle = GCHandle.Alloc(nonce, GCHandleType.Pinned);
                tagHandle = GCHandle.Alloc(tag, GCHandleType.Pinned);

                var info = new BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO
                {
                    cbSize = (uint)Marshal.SizeOf(typeof(BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO)),
                    dwInfoVersion = BCRYPT_AUTHENTICATED_CIPHER_MODE_INFO_VERSION,
                    pbNonce = nonceHandle.AddrOfPinnedObject(),
                    cbNonce = (uint)nonce.Length,
                    pbTag = tagHandle.AddrOfPinnedObject(),
                    cbTag = (uint)tag.Length
                };

                Check(BCryptDecrypt(hKey, ciphertext, ciphertext.Length, ref info, null, 0, plaintext, plaintext.Length, out _, 0));
                return plaintext;
            }
            finally
            {
                if (nonceHandle.IsAllocated) nonceHandle.Free();
                if (tagHandle.IsAllocated) tagHandle.Free();
                if (hKey != IntPtr.Zero) BCryptDestroyKey(hKey);
                if (hAlg != IntPtr.Zero) BCryptCloseAlgorithmProvider(hAlg, 0);
            }
        }

        private static void Check(uint status)
        {
            if (status != STATUS_SUCCESS)
                throw new CryptographicException($"BCrypt error: 0x{status:X8}");
        }
    }
}
