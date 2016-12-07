using Microsoft.Lync.Model;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ImgDataModel
{
    public static class Lync
    {

        private static LyncClient lyncClient;


        public static void init()
        {
            try
            {
                lyncClient = LyncClient.GetClient();
            }
            catch (ClientNotFoundException clientNotFoundException)
            {
                Console.WriteLine(clientNotFoundException);
                return;
            }
            catch (NotStartedByUserException notStartedByUserException)
            {
                Console.Out.WriteLine(notStartedByUserException);
                return;
            }
            catch (LyncClientException lyncClientException)
            {
                Console.Out.WriteLine(lyncClientException);
                return;
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                    return;
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

        }

        public static bool isLoggedIn()
        {
            try
            {
                if (lyncClient.State == ClientState.SignedIn)
                {
                    //Sign out If the current client state is Signed In

                    Console.WriteLine("User is Logged: " + lyncClient.Self);
                    return true;
                    //  lyncClient.BeginSignOut(SignOutCallback, null);

                }
                else if (lyncClient.State == ClientState.SignedOut)
                {
                    //Sign in If the current client state is Signed Out
                    return false;

                    // lyncClient.BeginSignIn(null, null, null, SignInCallback, null);

                }
            }
            catch (LyncClientException lyncClientException)
            {
                Console.WriteLine(lyncClientException);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }

            return false;
        }

        private static void SignInCallback(IAsyncResult result)
        {
            try
            {
                lyncClient.EndSignIn(result);
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
        }

        private static void SignOutCallback(IAsyncResult result)
        {
            try
            {
                lyncClient.EndSignOut(result);
            }
            catch (LyncClientException e)
            {
                Console.WriteLine(e);
            }
            catch (SystemException systemException)
            {
                if (IsLyncException(systemException))
                {
                    // Log the exception thrown by the Lync Model API.
                    Console.WriteLine("Error: " + systemException);
                }
                else
                {
                    // Rethrow the SystemException which did not come from the Lync Model API.
                    throw;
                }
            }
        }

        private static bool IsLyncException(SystemException ex)
        {
            return
                ex is NotImplementedException ||
                ex is ArgumentException ||
                ex is NullReferenceException ||
                ex is NotSupportedException ||
                ex is ArgumentOutOfRangeException ||
                ex is IndexOutOfRangeException ||
                ex is InvalidOperationException ||
                ex is TypeLoadException ||
                ex is TypeInitializationException ||
                ex is InvalidComObjectException ||
                ex is InvalidCastException;
        }
    }
}
