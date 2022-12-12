using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.DirectoryServices;

namespace WinDevOps.Accounts
{
    public class WindowsGroup
    {
        public readonly string PSComputerName;
        public string Name { get; internal set; }
        public string Description { get; internal set; }

        public List<WindowsUser> Users { get; internal set; }
        internal WindowsGroup(string computerName)
        {
            this.PSComputerName = string.IsNullOrEmpty(computerName) ? Environment.MachineName : computerName;
            Users = new List<WindowsUser>();
        }

        public override string ToString()
        {
            return $"Group: {Name}, Users: {string.Join(",", Users.Select(u=>u.Name).ToArray())}";
        }
    }

    public class WindowsUser
    {
        public string Name { get; internal set; }
        public string Domain { get; internal set; }
        public string Class { get; internal set; }
        public bool IsDomainUser { get; internal set; }

    }

    public static class AccountUtils
    {
        public static List<WindowsGroup> GetLocalWinUsers(string computerName = null, NetworkCredential credential = null)
        {
            var lst = new List<WindowsGroup>();
            var compName = string.IsNullOrEmpty(computerName) ? Environment.MachineName : computerName;
            var scope = WMI.WMIUtils.GetScope(computerName, credential);
            var groups = WMI.WMIUtils.PerformQuery(scope, "Select * FROM Win32_Group WHERE LocalAccount=True");
            foreach (var group in groups)
            {
                try
                {
                    DirectoryEntry de =
                        new DirectoryEntry($"WinNT://{compName}/{group.GetPropertyValue("Name").ToString()},group");
                    WindowsGroup winGroup = new WindowsGroup(compName);
                    winGroup.Description = de.InvokeGet("Description").ToString();
                    winGroup.Name = de.Name;
                    Console.WriteLine($"{compName}: Processing windows group '{winGroup.Name}'");
                    var members = (IEnumerable) de.Invoke("Members");
                    foreach (var member in members)
                    {
                        DirectoryEntry user = new DirectoryEntry(member);
                        WindowsUser winUser = new WindowsUser();
                        try
                        {
                            var path = user.InvokeGet("Adspath").ToString()
                                .Split(new[] {'/'}, StringSplitOptions.RemoveEmptyEntries);
                            var userClass = user.InvokeGet("Class").ToString();
                            winUser.Class = userClass.ToString();
                            winUser.Name = path[path.Length - 1];
                            winUser.Domain = path[path.Length - 2];
                            Console.WriteLine(
                                $"{compName}: Group '{winGroup.Name}' member '{winUser.Domain}\\{winUser.Name}'");
                            winUser.IsDomainUser = path.Length == 3;
                            winGroup.Users.Add(winUser);
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine($"{compName}: Failed to process group '{winGroup.Name}' member '{user.Name}'. Exception:{ex.Message}");
                        }
                    }

                    lst.Add(winGroup);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"{compName}: Failed to process group '{group.GetPropertyValue("Name")}'. Exception:{ex.Message}");
                }
            }

            return lst;
        }
    }
}
