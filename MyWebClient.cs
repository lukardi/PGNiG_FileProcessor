using CredentialManagement;
using System;

namespace PGNiG_FileProcessor
{

    public class MyWebClient
    {
        public string GetPassword(string KeyPair)
        {
            try
            {
                using (var cred = new Credential())
                {
                    cred.Target = KeyPair;
                    cred.Load();
                    return cred.Password;
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
            return "";
        }
        public string GetUsername(string KeyPair)
        {
            try
            {
                using (var cred = new Credential())
                {
                    cred.Target = KeyPair;
                    cred.Load();
                    return cred.Username;
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex);
            }
            return "";
        }

    }
}
