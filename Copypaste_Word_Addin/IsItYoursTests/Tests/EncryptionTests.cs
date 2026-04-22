using System;
using System.Linq;
using System.Collections.Generic;
using IsItYoursWordAddIn;

namespace IsItYoursTests.Tests
{
    static class EncryptionTests
    {
        public static IEnumerable<(string, Action)> All() => new (string, Action)[]
        {
            ("AES-GCM round-trip: decrypt returns original plaintext", AesGcmRoundTrip),
            ("AES-GCM: tampered ciphertext throws", AesGcmTamperDetect),
            ("AES-GCM: tampered tag throws", AesGcmTagTamperDetect),
            ("HMAC chain root: non-empty, 32 bytes", HmacChainRootNonEmpty),
            ("HMAC chain root: deterministic for same inputs", HmacChainRootDeterministic),
            ("HMAC chain root: different sessions produce different roots", HmacChainRootDifferentSessions),
            ("HMAC chain step: advances from root", HmacChainStepAdvances),
            ("HMAC chain step: different tick produces different HMAC", HmacChainStepDifferentTick),
            ("DPAPI wrap/unwrap round-trip", DpapiRoundTrip),
            ("DPAPI: wrong docGuid fails to unwrap", DpapiWrongGuid),
            ("RSA wrap: produces non-empty output", RsaWrapNonEmpty),
            ("GenerateAesKey: 32 bytes, random", GenerateAesKeyRandom),
        };

        static void AesGcmRoundTrip()
        {
            var key = TraceEncryption.GenerateAesKey();
            const string plaintext = "Hello, IsItYours! <tick id=\"abc12\" op=\"ins\" loc=\"0\" len=\"25\" paste=\"1\"/>";
            var (iv, ct, tag) = TraceEncryption.Encrypt(key, plaintext);
            var decrypted = TraceEncryption.Decrypt(key, iv, ct, tag);
            TestRunner.AssertEqual(plaintext, decrypted, "decrypted plaintext");
        }

        static void AesGcmTamperDetect()
        {
            var key = TraceEncryption.GenerateAesKey();
            var (iv, ct, tag) = TraceEncryption.Encrypt(key, "sensitive data");

            // Flip a byte in the ciphertext
            var ctBytes = Convert.FromBase64String(ct);
            ctBytes[0] ^= 0xFF;
            var tamperedCt = Convert.ToBase64String(ctBytes);

            bool threw = false;
            try { TraceEncryption.Decrypt(key, iv, tamperedCt, tag); }
            catch { threw = true; }
            TestRunner.Assert(threw, "tampered ciphertext must throw");
        }

        static void AesGcmTagTamperDetect()
        {
            var key = TraceEncryption.GenerateAesKey();
            var (iv, ct, tag) = TraceEncryption.Encrypt(key, "sensitive data");

            var tagBytes = Convert.FromBase64String(tag);
            tagBytes[0] ^= 0x01;
            var tamperedTag = Convert.ToBase64String(tagBytes);

            bool threw = false;
            try { TraceEncryption.Decrypt(key, iv, ct, tamperedTag); }
            catch { threw = true; }
            TestRunner.Assert(threw, "tampered tag must throw");
        }

        static void HmacChainRootNonEmpty()
        {
            var key = TraceEncryption.GenerateAesKey();
            var root = TraceEncryption.ChainRoot(key, "doc-guid-001", 0, new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc));
            TestRunner.Assert(root != null && root.Length == 32, "chain root is 32 bytes");
        }

        static void HmacChainRootDeterministic()
        {
            var key = new byte[32]; // all-zero key for determinism
            var t = new DateTime(2025, 6, 15, 12, 0, 0, DateTimeKind.Utc);
            var r1 = TraceEncryption.ChainRoot(key, "guid-abc", 5, t);
            var r2 = TraceEncryption.ChainRoot(key, "guid-abc", 5, t);
            TestRunner.Assert(r1.SequenceEqual(r2), "same inputs produce same root");
        }

        static void HmacChainRootDifferentSessions()
        {
            var key = new byte[32];
            var t = new DateTime(2025, 6, 15, 12, 0, 0, DateTimeKind.Utc);
            var r1 = TraceEncryption.ChainRoot(key, "guid-abc", 0, t);
            var r2 = TraceEncryption.ChainRoot(key, "guid-abc", 1, t);
            TestRunner.Assert(!r1.SequenceEqual(r2), "different session IDs produce different roots");
        }

        static void HmacChainStepAdvances()
        {
            var key = new byte[32];
            var t = new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var root = TraceEncryption.ChainRoot(key, "doc-x", 0, t);
            var step = TraceEncryption.ChainStep(key, root, "00000001a", "ins", 0, 25, 1);
            TestRunner.Assert(step != null && step.Length == 32, "step is 32 bytes");
            TestRunner.Assert(!step.SequenceEqual(root), "step differs from root");
        }

        static void HmacChainStepDifferentTick()
        {
            var key = new byte[32];
            var t = new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var root = TraceEncryption.ChainRoot(key, "doc-x", 0, t);
            var s1 = TraceEncryption.ChainStep(key, root, "00000001a", "ins", 0, 25, 1);
            var s2 = TraceEncryption.ChainStep(key, root, "00000001b", "ins", 0, 25, 1);
            TestRunner.Assert(!s1.SequenceEqual(s2), "different tickIds produce different steps");
        }

        static void DpapiRoundTrip()
        {
            var key = TraceEncryption.GenerateAesKey();
            const string docGuid = "test-doc-guid-12345";
            var wrapped = TraceEncryption.WrapAesKeyLocal(key, docGuid);
            TestRunner.Assert(!string.IsNullOrEmpty(wrapped), "wrapped key non-empty");
            var unwrapped = TraceEncryption.UnwrapAesKeyLocal(wrapped, docGuid);
            TestRunner.Assert(unwrapped != null, "unwrap succeeded");
            TestRunner.Assert(key.SequenceEqual(unwrapped), "unwrapped key matches original");
        }

        static void DpapiWrongGuid()
        {
            var key = TraceEncryption.GenerateAesKey();
            var wrapped = TraceEncryption.WrapAesKeyLocal(key, "correct-guid");
            var unwrapped = TraceEncryption.UnwrapAesKeyLocal(wrapped, "wrong-guid");
            // DPAPI with wrong entropy returns null (our catch returns null)
            TestRunner.Assert(unwrapped == null, "wrong docGuid must fail to unwrap");
        }

        static void RsaWrapNonEmpty()
        {
            var key = TraceEncryption.GenerateAesKey();
            var wrapped = TraceEncryption.WrapAesKey(key);
            TestRunner.Assert(!string.IsNullOrEmpty(wrapped), "RSA-wrapped key is non-empty");
            // RSA-2048 wraps to 256 bytes = 344 base64 chars
            var bytes = Convert.FromBase64String(wrapped);
            TestRunner.AssertEqual(256, bytes.Length, "RSA-2048 wrapped key length");
        }

        static void GenerateAesKeyRandom()
        {
            var k1 = TraceEncryption.GenerateAesKey();
            var k2 = TraceEncryption.GenerateAesKey();
            TestRunner.AssertEqual(32, k1.Length, "key length 32 bytes");
            TestRunner.Assert(!k1.SequenceEqual(k2), "two generated keys are different");
        }
    }
}
