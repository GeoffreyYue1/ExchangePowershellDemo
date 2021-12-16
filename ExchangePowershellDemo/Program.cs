using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.Security;

namespace ConsoleApp1
{
    class Program
    {


        static void Main(string[] args)
        {

            //RemotePowershell();
            CallExoV2();

            Console.ReadKey();

        }

        private static void CallExoV2()
        {


            using (PowerShell powerShell = PowerShell.Create())
            {
                // Source functions.
                powerShell.AddScript("Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force");
                powerShell.AddScript("Install-PackageProvider -Name NuGet  -Force");
                powerShell.AddScript("Install-Module -Name ExchangeOnlineManagement   -Force");
                powerShell.AddScript("Import-Module ExchangeOnlineManagement  -Force");
                powerShell.AddScript("$securepwd = ConvertTo-SecureString -String \"1234\" -AsPlainText -Force");
                powerShell.AddScript("Connect-ExchangeOnline -CertificateFilePath \"C:\\Work\\ProjTemp\\CallExPowershell\\mycert.pfx\" -CertificatePassword $securepwd -AppID \"a0b755fd-1050-406a-954c-428786b1bd67\" -Organization \"geoffrey1.onmicrosoft.com\"");

                // invoke execution on the pipeline (collecting output)
                Collection<PSObject> PSOutput = powerShell.Invoke();


                //powerShell.AddScript("Get-Mailbox");
                //PSOutput = powerShell.Invoke();

                //// loop through each output object item
                //foreach (PSObject outputItem in PSOutput)
                //{
                //    // if null object was dumped to the pipeline during the script then a null object may be present here
                //    if (outputItem != null)
                //    {
                //        Console.WriteLine($"Output line: [{outputItem}]");
                //    }
                //}

                GetMessageTrace(powerShell);

                // check the other output streams (for example, the error stream)
                if (powerShell.Streams.Error.Count > 0)
                {
                    // error records were written to the error stream.
                    // Do something with the error
                }


                powerShell.AddScript("Disconnect-ExchangeOnline -Confirm");

               PSOutput = powerShell.Invoke();

            }

        }

        private static void GetMessageTrace(PowerShell powershell)
        {
            powershell.AddScript("Get-MessageTrace -SenderAddress u10@geoffrey1.onmicrosoft.com -StartDate 12/07/2021 -EndDate 12/16/2021");

            Collection<PSObject> PSOutput = powershell.Invoke();

            

            // loop through each output object item
            foreach (PSObject outputItem in PSOutput)
            {
                // if null object was dumped to the pipeline during the script then a null object may be present here
                if (outputItem != null)
                {
                    Console.WriteLine($"Output line: SenderAddress - [{outputItem.Properties["SenderAddress"].Value}]; RecipientAddress - [{outputItem.Properties["RecipientAddress"].Value}]; Subject - [{outputItem.Properties["Subject"].Value}]");
                }
            }

        }

        private static void RemotePowershell()
        {
            string pwd = "Password1!";

            SecureString securePassword = new SecureString();
            //char c;
            foreach (char c in pwd.ToCharArray())
            {
                securePassword.AppendChar(c);
            }

            PSCredential cred = new PSCredential("admin001", securePassword);
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("http://ex1601.geoffrey.msftonlinelab.com/powershell?serializationLevel=Full"),
                                                                             "http://schemas.microsoft.com/powershell/Microsoft.Exchange", cred);

            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;


            GetMessageTraceReportDetails(connectionInfo);



        }


        private static void GetMessageTraceReportDetails(WSManConnectionInfo connectionInfo)
        {

        }


        private static void NewDistributionGroup(WSManConnectionInfo connectionInfo,string groupName)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
            PowerShell powershell = PowerShell.Create();

            string cmd1 = "New-DistributionGroup";
            powershell.AddCommand(cmd1);

            powershell.AddParameter("Name", groupName);


            runspace.Open();
            powershell.Runspace = runspace;
            Collection<PSObject> psResults = powershell.Invoke();

           if( psResults.Count>0)
            {
                Console.WriteLine(groupName + " Created");
            }
           else
            {
                Console.WriteLine(groupName + " Create failed");
            }

           

            runspace.Dispose();
            runspace = null;

            powershell.Dispose();
            powershell = null;
        }
        //public static string GetPropertyValue(this PSObject psObject, string propertyName)
        //{
        //    string ret = string.Empty;
        //    if (psObject.Properties[propertyName].Value != null)
        //        ret = psObject.Properties[propertyName].Value.ToString();

        //    return ret;
        //}

        private static void GetDistributionGroup(WSManConnectionInfo connectionInfo,string groupName)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
            PowerShell powershell = PowerShell.Create();

            string cmd1 = "Get-DistributionGroup";
            powershell.AddCommand(cmd1);

            powershell.AddParameter("Identity", groupName);


            runspace.Open();
            powershell.Runspace = runspace;
            Collection<PSObject> psResults = powershell.Invoke();

            if( psResults.Count>0)
            {
                Console.WriteLine(groupName + " successfully get");
            }
            else
            {
                Console.WriteLine(groupName + " faled to get");
            }

           

            runspace.Dispose();
            runspace = null;

            powershell.Dispose();
            powershell = null;
        }



    }
}
