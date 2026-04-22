// TestRunner.cs — minimal test harness (no NUnit dependency)
using System;
using System.Collections.Generic;

namespace IsItYoursTests
{
    class TestRunner
    {
        static int _pass, _fail;

        static void Main(string[] args)
        {
            Console.WriteLine("=== IsItYours Engine Tests ===\n");

            RunSuite("BTree", Tests.BTreeTests.All());
            RunSuite("Engine", Tests.EngineTests.All());
            RunSuite("Encryption", Tests.EncryptionTests.All());

            Console.WriteLine($"\n=== {_pass} passed, {_fail} failed ===");
            Console.WriteLine("\nNote: Integration tests for provenance chain require Word to be installed.");
            Console.WriteLine("See INTEGRATION_TEST_GUIDE.md for setup instructions.");
            Environment.Exit(_fail > 0 ? 1 : 0);
        }

        static void RunSuite(string name, IEnumerable<(string label, Action test)> tests)
        {
            Console.WriteLine($"[{name}]");
            foreach (var (label, test) in tests)
            {
                try
                {
                    test();
                    Console.WriteLine($"  PASS  {label}");
                    _pass++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  FAIL  {label}");
                    Console.WriteLine($"        {ex.Message}");
                    _fail++;
                }
            }
        }

        public static void Assert(bool condition, string message = "assertion failed")
        {
            if (!condition) throw new Exception(message);
        }

        public static void AssertEqual<T>(T expected, T actual, string label = "")
        {
            if (!Equals(expected, actual))
                throw new Exception($"{label} expected={expected} actual={actual}");
        }
    }
}
