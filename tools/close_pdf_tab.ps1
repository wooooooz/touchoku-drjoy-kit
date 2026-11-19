# close_pdf_tab.ps1 - close only the viewer tab that holds a public PDF open. 20261009
#
# the administrator 2026-10-09: the public duty-roster PDF was left open in PDF-XChange Editor, so the 15:00 run could not
# replace it. He never edits it, so the 15:00 run may close it - but only that tab (PDF-XChange holds ~60 other tabs).
#   1. Restart Manager (rstrtmgr.dll) tells which processes hold each file open.
#   2. In those processes' windows, UI Automation finds the document tab whose name is the file name without
#      ".pdf" (PDF-XChange shows the file name for our PDFs) and invokes its child button "タブを閉じる" (close tab).
#      Only that tab closes; nothing else in the viewer is touched. No keyboard, mouse or focus change.
#   3. Wait until the file is no longer held (up to -WaitSeconds).
# Output: one JSON line per file {file, held_by:[names], closed:n, released:bool, note}.
# Run with pwsh 7 (UIAutomationClient loads there on Windows). Called by dtx14_daily_sync.release_pdf().
#
#   pwsh -NoProfile -File tools\close_pdf_tab.ps1 -Path "<...>\R8年10月 病院当直勤務表.pdf" [-WhatIf]
#Requires -Version 7
param(
    [Parameter(Mandatory)][string[]]$Path,
    [int]$WaitSeconds = 15,
    [switch]$WhatIf
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices; using System.Collections.Generic;
public static class TouchokuRm {
 [StructLayout(LayoutKind.Sequential)] struct RM_UNIQUE_PROCESS { public int dwProcessId; public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime; }
 [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] struct RM_PROCESS_INFO { public RM_UNIQUE_PROCESS Process;
  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=256)] public string strAppName; [MarshalAs(UnmanagedType.ByValTStr, SizeConst=64)] public string strServiceShortName;
  public int ApplicationType; public uint AppStatus; public uint TSSessionId; [MarshalAs(UnmanagedType.Bool)] public bool bRestartable; }
 [DllImport("rstrtmgr.dll", CharSet=CharSet.Unicode)] static extern int RmStartSession(out uint h, int f, string key);
 [DllImport("rstrtmgr.dll")] static extern int RmEndSession(uint h);
 [DllImport("rstrtmgr.dll", CharSet=CharSet.Unicode)] static extern int RmRegisterResources(uint h, uint n, string[] files, uint a, RM_UNIQUE_PROCESS[] b, uint c, string[] d);
 [DllImport("rstrtmgr.dll")] static extern int RmGetList(uint h, out uint needed, ref uint n, [In, Out] RM_PROCESS_INFO[] info, ref uint reasons);
 public static List<KeyValuePair<int,string>> Holders(string path) {
  var r = new List<KeyValuePair<int,string>>(); uint h;
  if (RmStartSession(out h, 0, Guid.NewGuid().ToString()) != 0) return r;
  try {
   if (RmRegisterResources(h, 1, new[]{path}, 0, null, 0, null) != 0) return r;
   uint need = 0, n = 0, why = 0; RmGetList(h, out need, ref n, null, ref why);
   if (need == 0) return r;
   var arr = new RM_PROCESS_INFO[need]; n = need;
   if (RmGetList(h, out need, ref n, arr, ref why) == 0)
    for (int i = 0; i < n; i++) r.Add(new KeyValuePair<int,string>(arr[i].Process.dwProcessId, arr[i].strAppName));
  } finally { RmEndSession(h); }
  return r; } }
'@
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$UIA = [Windows.Automation.AutomationElement]
$CLOSE_NAMES = @('タブを閉じる', 'Close Tab', 'Close tab', '閉じる', 'Close')

function Close-Tabs([int]$ProcId, [string]$TabName) {
    $n = 0
    $wins = $UIA::RootElement.FindAll([Windows.Automation.TreeScope]::Children,
        (New-Object Windows.Automation.PropertyCondition($UIA::ProcessIdProperty, $ProcId)))
    foreach ($w in $wins) {
        $tabs = $w.FindAll([Windows.Automation.TreeScope]::Descendants,
            (New-Object Windows.Automation.PropertyCondition($UIA::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)))
        foreach ($t in $tabs) {
            $name = $t.Current.Name
            if ($name -ne $TabName -and $name -ne "$TabName.pdf") { continue }
            $btns = $t.FindAll([Windows.Automation.TreeScope]::Descendants,
                (New-Object Windows.Automation.PropertyCondition($UIA::ControlTypeProperty, [Windows.Automation.ControlType]::Button)))
            foreach ($b in $btns) {
                if ($CLOSE_NAMES -notcontains $b.Current.Name) { continue }
                if (-not $WhatIf) {
                    $b.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
                }
                $n++
                break
            }
        }
    }
    return $n
}

foreach ($f in $Path) {
    $res = [ordered]@{ file = $f; held_by = @(); closed = 0; released = $false; note = '' }
    try {
        $holders = [TouchokuRm]::Holders($f)
        $res.held_by = @($holders | ForEach-Object { "$($_.Value) ($($_.Key))" })
        if (-not $holders.Count) { $res.released = $true; $res.note = 'not held'; $res | ConvertTo-Json -Compress; continue }
        $stem = [IO.Path]::GetFileNameWithoutExtension($f)
        foreach ($h in $holders) { $res.closed += (Close-Tabs $h.Key $stem) }
        if ($WhatIf) { $res.note = 'whatif'; $res | ConvertTo-Json -Compress; continue }
        if (-not $res.closed) { $res.note = 'no matching tab'; $res | ConvertTo-Json -Compress; continue }
        $deadline = (Get-Date).AddSeconds($WaitSeconds)
        do {
            Start-Sleep -Milliseconds 500
            $left = [TouchokuRm]::Holders($f)
        } while ($left.Count -and (Get-Date) -lt $deadline)
        $res.released = -not $left.Count
        if ($left.Count) { $res.note = 'still held (a save prompt may be open)' }
    } catch {
        $res.note = "error: $($_.Exception.Message)"
    }
    $res | ConvertTo-Json -Compress
}
