using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text.Json;

using IPConfig.Languages;
using IPConfig.Models;

using Microsoft.Win32;

namespace IPConfig.Helpers;

/// <summary>
/// 管理网络配置信息的首选项。
/// </summary>
public static class NetworkManagement
{
    private static readonly int[] _gatewayCostMetric1 = [1];

    private static readonly Lazy<Dictionary<string, bool>> _physicalAdapters = new(GetPhysicalAdapters);

    public static NetworkInterface? GetActiveNetworkInterface()
    {
        var networks = NetworkInterface.GetAllNetworkInterfaces();

        var activeAdapter = networks.FirstOrDefault(x => x.NetworkInterfaceType != NetworkInterfaceType.Loopback
                            && x.NetworkInterfaceType != NetworkInterfaceType.Tunnel
                            && x.OperationalStatus == OperationalStatus.Up
                            && x.Name.StartsWith("vEthernet") == false);

        return activeAdapter;
    }

    public static IPAddress? GetIPAddress(string nicId)
    {
        foreach (var adapter in NetworkInterface.GetAllNetworkInterfaces())
        {
            if (adapter.Id == nicId)
            {
                foreach (var unicastIPAddressInformation in adapter.GetIPProperties().UnicastAddresses)
                {
                    if (unicastIPAddressInformation.Address.AddressFamily == AddressFamily.InterNetwork)
                    {
                        return unicastIPAddressInformation.Address;
                    }
                }
            }
        }

        return null;
    }

    public static IPInterfaceProperties? GetIPInterfaceProperties(string nicId)
    {
        foreach (var adapter in NetworkInterface.GetAllNetworkInterfaces())
        {
            if (adapter.Id == nicId)
            {
                return adapter.GetIPProperties();
            }
        }

        return null;
    }

    public static IPv4AdvancedConfig GetIPv4AdvancedConfig(NetworkInterface nic)
    {
        var config = GetIPAdvancedConfig(nic, AddressFamily.InterNetwork);

        return (IPv4AdvancedConfig)config;
    }

    public static IPv4Config GetIPv4Config(NetworkInterface nic)
    {
        var props = nic.GetIPProperties();

        var info = props.UnicastAddresses
            .FirstOrDefault(x => x.Address.AddressFamily == AddressFamily.InterNetwork);

        bool isDhcpEnabled = IsDhcpEnabled(nic.Id);

        string ip = info?.Address.ToString() ?? "";
        string mask = info?.IPv4Mask.ToString() ?? "";
        string gateway = props.GatewayAddresses.FirstOrDefault()?.Address.ToString() ?? "";

        bool isAutoDns = IsAutoDns(nic.Id);

        string dns1 = "";
        string dns2 = "";

        if (!isAutoDns)
        {
            if (props.DnsAddresses is [var pref, var alt, ..])
            {
                dns1 = pref.ToString();
                dns2 = alt.ToString();
            }
            else if (props.DnsAddresses is [var p])
            {
                dns1 = p.ToString();
            }
        }

        var cfg = new IPv4Config() {
            IsDhcpEnabled = isDhcpEnabled,
            IP = ip,
            Mask = mask,
            Gateway = gateway,
            IsAutoDns = isAutoDns,
            Dns1 = dns1,
            Dns2 = dns2
        };

        return cfg;
    }

    public static IPv6AdvancedConfig GetIPv6AdvancedConfig(NetworkInterface nic)
    {
        var config = GetIPAdvancedConfig(nic, AddressFamily.InterNetworkV6);

        return (IPv6AdvancedConfig)config;
    }

    public static bool IsAutoDns(string nicId)
    {
        string path = $@"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{nicId}";
        string? ns = Registry.GetValue(path, "NameServer", null) as string;

        // Wi-Fi 可能使用配置文件而不是网卡配置，需要优先判断。
        if (Registry.GetValue(path, "ProfileNameServer", null) is not string pns)
        {
            return String.IsNullOrEmpty(ns);
        }

        return String.IsNullOrEmpty(pns);
    }

    public static bool IsDhcpEnabled(string nicId)
    {
        string path = $@"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\{nicId}";
        object? dhcp = Registry.GetValue(path, "EnableDHCP", 0);

        return dhcp is 1;
    }

    public static bool IsPhysicalAdapter(string nicId)
    {
        if (!_physicalAdapters.IsValueCreated)
        {
            GetPhysicalAdapters();
        }

        return _physicalAdapters.Value.GetValueOrDefault(nicId);
    }

    public static void SetIPv4(string nicId, string ipAddress, string subnetMask, string gateway)
    {
        using var networkConfigMng = new ManagementClass("Win32_NetworkAdapterConfiguration");
        using var networkConfigs = networkConfigMng.GetInstances();

        using var managementObject = networkConfigs
            .OfType<ManagementObject>()
            .FirstOrDefault(objMO => (bool)objMO["IPEnabled"] && objMO["SettingID"].Equals(nicId));

        if (managementObject is null)
        {
            return;
        }

        using var newIP = managementObject.GetMethodParameters("EnableStatic");

        if ((!String.IsNullOrEmpty(ipAddress)) || (!String.IsNullOrEmpty(subnetMask)))
        {
            if (!String.IsNullOrEmpty(ipAddress))
            {
                ConcatManagementBaseObjectValues(newIP, "IPAddress", new[] { ipAddress });
            }

            if (!String.IsNullOrEmpty(subnetMask))
            {
                ConcatManagementBaseObjectValues(newIP, "SubnetMask", new[] { subnetMask });
            }

            managementObject.InvokeMethod("EnableStatic", newIP, null!);
        }

        if (!String.IsNullOrEmpty(gateway))
        {
            using var newGateway = managementObject.GetMethodParameters("SetGateways");

            ConcatManagementBaseObjectValues(newGateway, "DefaultIPGateway", new[] { gateway });
            ConcatManagementBaseObjectValues(newGateway, "GatewayCostMetric", _gatewayCostMetric1);
            managementObject.InvokeMethod("SetGateways", newGateway, null!);
        }
    }

    public static void SetIPv4Dhcp(string nicId)
    {
        using var networkConfigMng = new ManagementClass("Win32_NetworkAdapterConfiguration");
        using var networkConfigs = networkConfigMng.GetInstances();

        using var managementObject = networkConfigs
            .OfType<ManagementObject>()
            .FirstOrDefault(objMO => (bool)objMO["IPEnabled"] && objMO["SettingID"].Equals(nicId));

        if (managementObject is null)
        {
            return;
        }

        managementObject.InvokeMethod("SetDNSServerSearchOrder", null!);
        managementObject.InvokeMethod("EnableDHCP", null!);
    }

    public static void SetIPv4Dns(string nicId, string dns1, string dns2 = "")
    {
        using var networkConfigMng = new ManagementClass("Win32_NetworkAdapterConfiguration");
        using var networkConfigs = networkConfigMng.GetInstances();

        using var managementObject = networkConfigs
            .OfType<ManagementObject>()
            .FirstOrDefault(objMO => (bool)objMO["IPEnabled"] && objMO["SettingID"].Equals(nicId));

        if (managementObject is null)
        {
            return;
        }

        SetManagementObjectArrayValue(managementObject,
            "SetDNSServerSearchOrder",
            "DNSServerSearchOrder",
            new[] { dns1, dns2 });
    }

    public static void SetIPv4DnsAuto(string nicId)
    {
        using var networkConfigMng = new ManagementClass("Win32_NetworkAdapterConfiguration");
        using var networkConfigs = networkConfigMng.GetInstances();

        using var managementObject = networkConfigs
            .OfType<ManagementObject>()
            .FirstOrDefault(objMO => (bool)objMO["IPEnabled"] && objMO["SettingID"].Equals(nicId));

        if (managementObject is null)
        {
            return;
        }

        using var dnsObj = managementObject.GetMethodParameters("SetDNSServerSearchOrder");

        if (dnsObj is null)
        {
            return;
        }

        dnsObj["DNSServerSearchOrder"] = null!;
        managementObject.InvokeMethod("SetDNSServerSearchOrder", dnsObj, null!);
    }

    private static void ConcatManagementBaseObjectValues<T>(ManagementBaseObject objMBO, string propertyName, T[] newValues)
    {
        if (objMBO[propertyName] is T[] oldValues)
        {
            objMBO[propertyName] = newValues.Concat(oldValues.Skip(newValues.Length)).ToArray();
        }
        else
        {
            objMBO[propertyName] = newValues;
        }
    }

    private static IPAdvancedConfigBase GetIPAdvancedConfig(NetworkInterface nic, AddressFamily addressFamily)
    {
        var props = nic.GetIPProperties();

        var info = props.UnicastAddresses
            .Where(x => x.Address.AddressFamily == addressFamily);

        var iPCollection = info.Select(x => x.Address.ToString());
        var prefixLenghtCollection = info.Select(x => x.PrefixLength);

        var prefixOriginCollection = info.Select(x => x.PrefixOrigin);
        var suffixOriginCollection = info.Select(x => x.SuffixOrigin);

        var validLifetimeCollection = info.Select(x => {
            var when = (DateTime.UtcNow + TimeSpan.FromSeconds(x.AddressValidLifetime)).ToLocalTime();

            return when.ToString(LangSource.Instance.CurrentCulture);
        });

        var prefLifetimeCollection = info.Select(x => {
            var when = (DateTime.UtcNow + TimeSpan.FromSeconds(x.AddressPreferredLifetime)).ToLocalTime();

            return when.ToString(LangSource.Instance.CurrentCulture);
        });

        var dhcpLeaseLifetimeCollection = info.Select(x => {
            var when = (DateTime.UtcNow + TimeSpan.FromSeconds(x.DhcpLeaseLifetime)).ToLocalTime();

            return when.ToString(LangSource.Instance.CurrentCulture);
        });

        var isTransientCollection = info.Select(x => x.IsTransient);
        var dadsCollection = info.Select(x => x.DuplicateAddressDetectionState);
        var isDnsEligibleCollcetion = info.Select(x => x.IsDnsEligible);

        if (addressFamily == AddressFamily.InterNetwork)
        {
            bool isDhcpEnabled = IsDhcpEnabled(nic.Id);
            var maskCollection = info.Select(x => x.IPv4Mask.ToString());

            var gatewayCollection = props.GatewayAddresses
                .Where(x => x.Address.AddressFamily == AddressFamily.InterNetwork)
                .Select(x => x.Address.ToString());

            bool isAutoDns = IsAutoDns(nic.Id);

            var dnsCollection = props.DnsAddresses
                .Where(x => x.AddressFamily == AddressFamily.InterNetwork)
                .Select(x => x.ToString());

            var winsCollection = nic.GetIPProperties().WinsServersAddresses.Select(x => x.ToString());

            var iPv4AdvancedConfig = new IPv4AdvancedConfig() {
                IsDhcpEnabled = isDhcpEnabled,
                PreferredIP = iPCollection.FirstOrDefault(""),
                AlternateIPCollection = iPCollection.Skip(1).ToImmutableList(),
                PreferredMask = maskCollection.FirstOrDefault(""),
                AlternateMaskCollection = maskCollection.Skip(1).ToImmutableList(),
                PreferredGateway = gatewayCollection.FirstOrDefault(""),
                AlternateGatewayCollection = gatewayCollection.Skip(1).ToImmutableList(),
                IsAutoDns = isAutoDns,
                PreferredDns = dnsCollection.FirstOrDefault(""),
                AlternateDnsCollection = dnsCollection.Skip(1).ToImmutableList(),
                WinsServerAddress = winsCollection.FirstOrDefault(""),
                WinsServerAddressCollection = winsCollection.Skip(1).ToImmutableList(),
                ValidLifetime = validLifetimeCollection.FirstOrDefault(""),
                ValidLifetimeCollection = validLifetimeCollection.Skip(1).ToImmutableList(),
                PreferredLifetime = prefLifetimeCollection.FirstOrDefault(""),
                PreferredLifetimeCollection = prefLifetimeCollection.Skip(1).ToImmutableList(),
                DhcpLeaseLifetime = dhcpLeaseLifetimeCollection.FirstOrDefault(""),
                DhcpLeaseLifetimeCollection = dhcpLeaseLifetimeCollection.Skip(1).ToImmutableList(),
                IsTransient = isTransientCollection.FirstOrDefault(),
                IsTransientCollection = isTransientCollection.Skip(1).ToImmutableList(),
                IsDnsEligible = isDnsEligibleCollcetion.FirstOrDefault(),
                IsDnsEligibleCollcetion = isDnsEligibleCollcetion.Skip(1).ToImmutableList(),
                DuplicateAddressDetectionState = dadsCollection.FirstOrDefault(),
                DuplicateAddressDetectionStateCollcetion = dadsCollection.Skip(1).ToImmutableList()
            };

            return iPv4AdvancedConfig;
        }
        else if (addressFamily == AddressFamily.InterNetworkV6)
        {
            var gatewayCollection = props.GatewayAddresses
                .Where(x => x.Address.AddressFamily == AddressFamily.InterNetworkV6)
                .Select(x => x.Address.ToString());

            var dnsCollection = props.DnsAddresses
                .Where(x => x.AddressFamily == AddressFamily.InterNetworkV6)
                .Select(x => x.ToString());

            var config = new IPv6AdvancedConfig() {
                PreferredIP = iPCollection.FirstOrDefault(""),
                AlternateIPCollection = iPCollection.Skip(1).ToImmutableList(),
                PreferredPrefixLength = prefixLenghtCollection.FirstOrDefault(),
                AlternatePrefixLengthCollection = prefixLenghtCollection.Skip(1).ToImmutableList(),
                PreferredGateway = gatewayCollection.FirstOrDefault(""),
                AlternateGatewayCollection = gatewayCollection.Skip(1).ToImmutableList(),
                PrefixOrigin = prefixOriginCollection.FirstOrDefault(),
                PrefixOriginCollection = prefixOriginCollection.Skip(1).ToImmutableList(),
                SuffixOrigin = suffixOriginCollection.FirstOrDefault(),
                SuffixOriginCollection = suffixOriginCollection.Skip(1).ToImmutableList(),
                PreferredDns = dnsCollection.FirstOrDefault(""),
                AlternateDnsCollection = dnsCollection.Skip(1).ToImmutableList(),
                ValidLifetime = validLifetimeCollection.FirstOrDefault(""),
                ValidLifetimeCollection = validLifetimeCollection.Skip(1).ToImmutableList(),
                PreferredLifetime = prefLifetimeCollection.FirstOrDefault(""),
                PreferredLifetimeCollection = prefLifetimeCollection.Skip(1).ToImmutableList(),
                DhcpLeaseLifetime = dhcpLeaseLifetimeCollection.FirstOrDefault(""),
                DhcpLeaseLifetimeCollection = dhcpLeaseLifetimeCollection.Skip(1).ToImmutableList(),
                IsTransient = isTransientCollection.FirstOrDefault(),
                IsTransientCollection = isTransientCollection.Skip(1).ToImmutableList(),
                IsDnsEligible = isDnsEligibleCollcetion.FirstOrDefault(),
                IsDnsEligibleCollcetion = isDnsEligibleCollcetion.Skip(1).ToImmutableList(),
                DuplicateAddressDetectionState = dadsCollection.FirstOrDefault(),
                DuplicateAddressDetectionStateCollcetion = dadsCollection.Skip(1).ToImmutableList()
            };

            return config;
        }

        throw new ArgumentOutOfRangeException(nameof(addressFamily), addressFamily, $"Supports {AddressFamily.InterNetwork} and {AddressFamily.InterNetworkV6} only.");
    }


    public static InterfaceMetricInfo GetInterfaceMetricInfo(int interfaceIndex, AddressFamily addressFamily)
    {
        string familyName = addressFamily == AddressFamily.InterNetwork
            ? "IPv4"
            : addressFamily == AddressFamily.InterNetworkV6
                ? "IPv6"
                : throw new ArgumentOutOfRangeException(nameof(addressFamily), addressFamily, null);

        string script = $$"""
            $info = Get-NetIPInterface -InterfaceIndex {{interfaceIndex}} -AddressFamily {{familyName}} -ErrorAction Stop |
                Select-Object -First 1 @{Name='InterfaceMetric';Expression={[int]$_.InterfaceMetric}}, @{Name='AutomaticMetric';Expression={[string]$_.AutomaticMetric}}
            $info | ConvertTo-Json -Compress
            """;

        try
        {
            string json = RunPowerShell(script);

            if (String.IsNullOrWhiteSpace(json))
            {
                return new(null, null);
            }

            var result = JsonSerializer.Deserialize<PowerShellInterfaceMetricInfo>(json);

            if (result is null)
            {
                return new(null, null);
            }

            bool? automatic = result.AutomaticMetric?.Equals("Enabled", StringComparison.OrdinalIgnoreCase) switch {
                true => true,
                false => result.AutomaticMetric?.Equals("Disabled", StringComparison.OrdinalIgnoreCase) switch {
                    true => false,
                    false => null
                }
            };

            return new(result.InterfaceMetric, automatic);
        }
        catch
        {
            return new(null, null);
        }
    }

    public static void PreferConnectionType(ConnectionType preferredType)
    {
        foreach (var nic in GetPrioritizableNics())
        {
            int metric = nic.ConnectionType == preferredType ? 5 : 50;
            SetInterfaceMetric(nic, metric);
        }
    }

    public static void ResetConnectionPriority()
    {
        foreach (var nic in GetPrioritizableNics())
        {
            ResetInterfaceMetric(nic);
        }
    }


    private static IEnumerable<Nic> GetPrioritizableNics()
    {
        return NetworkInterface.GetAllNetworkInterfaces()
            .Select(x => new Nic(x))
            .Where(x => IsPhysicalAdapter(x.Id)
                        && (x.ConnectionType == ConnectionType.Ethernet || x.ConnectionType == ConnectionType.Wlan));
    }

    private static void ResetInterfaceMetric(Nic nic)
    {
        if (nic.SupportsIPv4 && nic.IPv4InterfaceProperties is not null)
        {
            SetAutomaticMetric(nic.IPv4InterfaceProperties.Index, AddressFamily.InterNetwork, true);
        }

        if (nic.SupportsIPv6 && nic.IPv6InterfaceProperties is not null)
        {
            SetAutomaticMetric(nic.IPv6InterfaceProperties.Index, AddressFamily.InterNetworkV6, true);
        }
    }

    private static void SetInterfaceMetric(Nic nic, int metric)
    {
        if (nic.SupportsIPv4 && nic.IPv4InterfaceProperties is not null)
        {
            SetAutomaticMetric(nic.IPv4InterfaceProperties.Index, AddressFamily.InterNetwork, false, metric);
        }

        if (nic.SupportsIPv6 && nic.IPv6InterfaceProperties is not null)
        {
            SetAutomaticMetric(nic.IPv6InterfaceProperties.Index, AddressFamily.InterNetworkV6, false, metric);
        }
    }

    private static void SetAutomaticMetric(int interfaceIndex, AddressFamily addressFamily, bool automatic, int? metric = null)
    {
        string familyName = addressFamily == AddressFamily.InterNetwork
            ? "IPv4"
            : addressFamily == AddressFamily.InterNetworkV6
                ? "IPv6"
                : throw new ArgumentOutOfRangeException(nameof(addressFamily), addressFamily, null);

        string script = automatic
            ? $$"Set-NetIPInterface -InterfaceIndex {{interfaceIndex}} -AddressFamily {{familyName}} -AutomaticMetric Enabled -ErrorAction Stop | Out-Null"
            : $$"Set-NetIPInterface -InterfaceIndex {{interfaceIndex}} -AddressFamily {{familyName}} -AutomaticMetric Disabled -InterfaceMetric {{metric}} -ErrorAction Stop | Out-Null";

        RunPowerShell(script);
    }

    private static string RunPowerShell(string script)
    {
        string encodedScript = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(script));

        var psi = new ProcessStartInfo {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encodedScript}",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi);

        if (process is null)
        {
            throw new InvalidOperationException("Failed to start PowerShell.");
        }

        string stdout = process.StandardOutput.ReadToEnd();
        string stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(String.IsNullOrWhiteSpace(stderr) ? stdout.Trim() : stderr.Trim());
        }

        return stdout.Trim();
    }

    private sealed class PowerShellInterfaceMetricInfo
    {
        public int? InterfaceMetric { get; set; }

        public string? AutomaticMetric { get; set; }
    }

    private static Dictionary<string, bool> GetPhysicalAdapters()
    {
        using var searcher = new ManagementObjectSearcher(@"root\CIMV2",
            $@"SELECT PhysicalAdapter, GUID FROM Win32_NetworkAdapter WHERE NOT PNPDeviceID LIKE 'ROOT\\%'");

        var managementObjects = searcher.Get().OfType<ManagementObject>();
        var result = new Dictionary<string, bool>();

        foreach (var objMO in managementObjects)
        {
            if (objMO?.Properties["GUID"].Value is string id)
            {
                bool isPhysical = Convert.ToBoolean(objMO?.Properties["PhysicalAdapter"].Value);
                result[id] = isPhysical;
            }
        }

        return result;
    }

    private static void SetManagementObjectArrayValue<T>(ManagementObject objMO, string methodName, string propertyName, T[] newValues)
    {
        using var objMBO = objMO.GetMethodParameters(methodName);

        if (objMBO is null)
        {
            return;
        }

        var oldValues = objMBO[propertyName] as T[] ?? [];
        objMBO[propertyName] = newValues.Concat(oldValues.Skip(newValues.Length)).ToArray();
        objMO.InvokeMethod(methodName, objMBO, null!);
    }
}
