# Security Review (fork audit)

This is a quick source audit of the upstream `CodingOctocat/IPConfig` project as imported into this fork.

## What was checked

- External network calls
- Process spawning and shell execution
- Registry and WMI usage
- File writes and local persistence
- Signs of autoruns, services, scheduled tasks, hidden downloads, or embedded script runners

## Findings

### Expected privileged behavior

The application uses WMI (`Win32_NetworkAdapterConfiguration`) and registry reads/writes related to Windows network adapter state. That is expected for a GUI tool that changes IP, gateway, DHCP, and DNS settings.

### External network activity

The app checks GitHub releases to display version/update information. No other obvious outbound HTTP endpoints were found in the source reviewed.

### Process execution

The original upstream project used `cmd /c start ...` to open URLs. In this fork that was replaced with direct shell execution (`UseShellExecute = true`) to avoid an unnecessary command-shell hop.

For the new interface-priority feature, the fork also uses `powershell.exe` to call the built-in `Get-NetIPInterface` and `Set-NetIPInterface` cmdlets so it can read, set, and reset interface metrics in a Windows-supported way.

### Local persistence

The app stores local data in a LiteDB file (`ipconfig.db`), writes backup files under `backup/`, and writes `error.log` when an unhandled exception is shown. These are normal local persistence behaviors for this app.

### Things not found

No evidence was found in the reviewed source for:

- scheduled tasks
- service installation
- autorun registration
- hidden download-and-execute flows
- PowerShell execution
- credential exfiltration logic
- extra network beacons beyond the GitHub release check

## Residual risk / recommendations

- Review NuGet dependency versions before releasing a production build.
- Consider making the GitHub release check optional if you want a quieter build.
- Build from source yourself and verify the produced binary hash before distribution.

## Conclusion

No obvious backdoor behavior was identified in the source reviewed for this fork. The main hardening change applied here is the removal of the `cmd /c start` URL-launch pattern.
