using System;

namespace SnipperCloneCleanFinal.Infrastructure
{
    public static class AuthManager
    {
        public static bool IsAuthenticated { get; private set; } = true;

        public static bool Authenticate(string username, string password)
        {
            // Simple authentication placeholder
            IsAuthenticated = true;
            return true;
        }

        public static void Logout()
        {
            IsAuthenticated = false;
        }
    }
} 