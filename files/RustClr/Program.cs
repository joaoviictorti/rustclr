using System;
using System.Runtime.InteropServices;

namespace RustClr
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("[CLR] Program running");

            // Show arguments
            if (args != null && args.Length > 0)
            {
                Console.WriteLine("[CLR] Args:");
                foreach (var a in args)
                    Console.WriteLine($"- {a}");
            }
            else
            {
                Console.WriteLine("[CLR] No args provided");
            }

            // Show loaded AppDomain
            Console.WriteLine($"[CLR] AppDomain: {AppDomain.CurrentDomain.FriendlyName}");

            // Show runtime information
            Console.WriteLine($"[CLR] Environment.Version: {Environment.Version}");
            Console.WriteLine($"[CLR] FrameworkDescription: {RuntimeInformation.FrameworkDescription}");

            try
            {
                var tf = AppDomain.CurrentDomain.SetupInformation.TargetFrameworkName;
                if (!string.IsNullOrEmpty(tf))
                    Console.WriteLine($"[CLR] TargetFramework: {tf}");
            }
            catch { }

            Console.WriteLine("[CLR] End");

            // Exit test (forces exit code 1337)
            Console.WriteLine("[CLR] Calling Environment.Exit(1337)...");
            Environment.Exit(1337);

            // If your patch is working, THIS LINE WILL PRINT.
            Console.WriteLine("[CLR] Exit was intercepted");
        }
    }
}
