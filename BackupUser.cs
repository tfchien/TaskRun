using System;

namespace TaskRun
{
    public class BackupUser
    {
        public string User { get; set; }
        public string Password { get; set; }
        public bool Encrypted { get; set; }

        public BackupUser() { }

        public BackupUser(string user, string password, bool encrypted)
        {
            User = user;
            Password = password;
            Encrypted = encrypted;
        }

        public override string ToString()
        {
            return $"User: {User}, Encrypted: {Encrypted}";
        }
    }
}
