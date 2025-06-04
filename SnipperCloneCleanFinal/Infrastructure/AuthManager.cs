using System;

namespace SnipperCloneCleanFinal.Infrastructure
{
    public static class AuthManager
    {
        public static bool IsAuthenticated { get; private set; } = true;

        public static bool Authenticate(string username, string password)
        {
            // Simplified authentication - always return true for development
            IsAuthenticated = true;
            return IsAuthenticated;
        }

        public static void Logout()
        {
            IsAuthenticated = false;
        }
    }
} 