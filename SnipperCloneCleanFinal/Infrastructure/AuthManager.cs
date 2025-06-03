using System;

namespace SnipperCloneCleanFinal.Infrastructure
{
    public static class AuthManager
    {
        public static bool IsAuthenticated { get; private set; } = false;

        public static bool Authenticate(string username, string password)
        {
            var validUser = Environment.GetEnvironmentVariable("SNIPPER_USER") ?? "admin";
            var validPass = Environment.GetEnvironmentVariable("SNIPPER_PASS") ?? "snipper";
            IsAuthenticated = string.Equals(username, validUser, StringComparison.OrdinalIgnoreCase)
                && password == validPass;
            return IsAuthenticated;
        }

        public static void Logout()
        {
            IsAuthenticated = false;
        }
    }
} 