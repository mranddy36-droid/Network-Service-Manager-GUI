"""Python GUI port of Network Service Manager (BESKAR) v2.5.4.

Windows 10/11. Requires Python 3.10+ and PowerShell; administrator rights are
needed only for network changes. Profile files remain compatible with the PS version.
"""
from __future__ import annotations

import base64
import ctypes
import hashlib
import ipaddress
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, simpledialog, ttk
from typing import Any

APP_DIR = Path(__file__).resolve().parent
PROFILES_DIR = APP_DIR / "Profiles"
LOG_DIR = APP_DIR / "Logs"
SNAPSHOT = APP_DIR / "OriginalNetworkState.json"
SNAPSHOT_DIR = APP_DIR / "Snapshots"
SETTINGS = APP_DIR / "PythonSettings.json"
APP_NAME = "Network Service Manager"
VLAN_KEYS = ("VLAN_ID", "RegVlanID", "VlanID", "VLANID", "*VlanID")
MODE_LABELS = {
    "Network": "Один IP-адрес",
    "MultiAddress": "Несколько IP-адресов",
    "Routes": "Только маршруты",
    "Mixed": "IP-адреса и маршруты",
    "DHCP": "Автоматически (DHCP)",
}
MODE_DESCRIPTIONS = {
    "Network": "Назначает один статический IPv4-адрес и маску. Можно указать шлюз и VLAN. Подходит для обычного подключения к одному устройству или сети.",
    "MultiAddress": "Назначает несколько IPv4-адресов одному адаптеру. Используйте, если к устройству нужно обращаться из нескольких подсетей. Можно указать VLAN.",
    "Routes": "Добавляет маршруты до удалённых подсетей через указанные шлюзы. IP-адрес адаптера не меняется. Подходит, когда адаптер уже настроен.",
    "Mixed": "Сначала назначает несколько IPv4-адресов, затем добавляет маршруты. Используйте, когда нужны и адреса для локальных подсетей, и пути к удалённым сетям. Можно указать VLAN.",
    "DHCP": "Автоматически получает IPv4-адрес, шлюз и DNS от DHCP-сервера. Можно задать VLAN и дополнительные маршруты. Ручные IPv4-поля не нужны.",
}


def mode_from_label(label: str) -> str:
    return next((key for key, value in MODE_LABELS.items() if value == label), label)


def setup_logging() -> None:
    LOG_DIR.mkdir(exist_ok=True)
    logging.basicConfig(filename=LOG_DIR / f"BESKAR_{datetime.now():%Y-%m}.log",
                        level=logging.INFO, encoding="utf-8",
                        format="%(asctime)s %(levelname)s %(message)s")


def run_powershell(script: str, timeout: int = 45) -> str:
    """Execute a PowerShell 5.1-compatible script and return stdout or raise."""
    # Windows PowerShell 5.1 may use the active OEM code page for redirected
    # stdout. Set UTF-8 inside the process before emitting JSON/text so adapter
    # aliases containing Cyrillic characters survive the subprocess boundary.
    script = ("[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
              "$OutputEncoding = [Console]::OutputEncoding; " + script)
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    exe = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                       "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
    result = subprocess.run([exe, "-NoLogo", "-NoProfile", "-NonInteractive",
                             "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
                            capture_output=True, text=True, timeout=timeout,
                            encoding="utf-8", errors="replace", creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if result.returncode:
        raise RuntimeError((result.stderr or result.stdout or "PowerShell failed").strip())
    return result.stdout.strip()


def ps_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def is_administrator() -> bool:
    """Return whether Windows started this process with elevated rights."""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except (AttributeError, OSError):
        return False


def user_friendly_error(error: Exception) -> str:
    """Keep PowerShell diagnostic output in the log, not in a GUI dialog."""
    details = str(error).strip()
    lower = details.lower()
    if "set-netadapteradvancedproperty" in lower:
        if "permissiondenied" in lower or "access is denied" in lower or "access denied" in lower:
            return ("Не удалось изменить VLAN: Windows требует права администратора.\n\n"
                    "Закройте программу и запустите Start-Python.bat через «Запуск от имени администратора».")
        return ("Не удалось изменить VLAN в драйвере сетевого адаптера. "
                "Проверьте поддержку VLAN драйвером и права администратора.")
    if "permissiondenied" in lower or "access is denied" in lower or "access denied" in lower:
        return ("Windows отклонила изменение сетевых параметров.\n\n"
                "Закройте программу и запустите Start-Python.bat через «Запуск от имени администратора».")
    if "#< clixml" in lower:
        return "Не удалось применить сетевые параметры. Технические сведения сохранены в журнале Logs."
    return details or "Не удалось выполнить операцию. Технические сведения сохранены в журнале Logs."


def mask_from(value: str) -> str:
    value = value.strip()
    if value.startswith("/"):
        value = value[1:]
    if value.isdigit():
        n = int(value)
        if not 0 <= n <= 32:
            raise ValueError("Префикс маски должен быть от 0 до 32")
        return str(ipaddress.IPv4Network(f"0.0.0.0/{n}").netmask)
    return str(ipaddress.IPv4Network(f"0.0.0.0/{value}").netmask)


def validate_profile(p: dict[str, Any]) -> None:
    if not isinstance(p, dict) or not str(p.get("Name", "")).strip():
        raise ValueError("В профиле должно быть непустое поле Name")
    mode = p.get("Mode") or ("Mixed" if "Addresses" in p and "Routes" in p else
                              "MultiAddress" if "Addresses" in p else "Routes" if "Routes" in p else "Network")
    if mode not in ("Network", "MultiAddress", "Routes", "Mixed", "DHCP"):
        raise ValueError("Неизвестный тип профиля")
    p["Mode"] = mode
    if p.get("VLAN") not in (None, ""):
        p["VLAN"] = int(p["VLAN"])
        if not 1 <= p["VLAN"] <= 4094:
            raise ValueError("VLAN должен быть от 1 до 4094")
    if mode == "Network":
        ipaddress.IPv4Address(p["IP"])
        p["Mask"] = mask_from(str(p.get("Mask", "24")))
        if p.get("Gateway"):
            ipaddress.IPv4Address(p["Gateway"])
    if mode in ("MultiAddress", "Mixed"):
        if not p.get("Addresses"):
            raise ValueError("Добавьте хотя бы один IPv4-адрес")
        for a in p.get("Addresses", []):
            ipaddress.IPv4Address(a["IP"])
            a["Mask"] = mask_from(str(a.get("Mask", "24")))
    if mode in ("Routes", "Mixed", "DHCP"):
        if mode != "DHCP" and not p.get("Routes"):
            raise ValueError("Добавьте хотя бы один маршрут")
        for r in p.get("Routes", []):
            ipaddress.IPv4Address(r["Destination"])
            ipaddress.IPv4Address(r["Gateway"])
            r["Mask"] = mask_from(str(r["Mask"]))
    for server in p.get("DnsServers", []):
        ipaddress.IPv4Address(server)


class NetworkBackend:
    """PowerShell/netsh adapter backend; all dynamic values are quoted safely."""
    def adapters(self) -> list[dict[str, Any]]:
        script = """$ErrorActionPreference='Stop'; Get-NetAdapter -Physical | Sort-Object Name | Select-Object Name,InterfaceGuid,InterfaceDescription,Status,LinkSpeed | ConvertTo-Json -Depth 4 -Compress"""
        data = json.loads(run_powershell(script) or "[]")
        return data if isinstance(data, list) else [data]

    def state(self, name: str) -> dict[str, Any]:
        n = ps_quote(name)
        script = f"""$ErrorActionPreference='Stop'; $a=Get-NetAdapter -Name {n}; $ip=Get-NetIPInterface -InterfaceAlias {n} -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1; $ad=@(Get-NetIPAddress -InterfaceAlias {n} -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {{$_.IPAddress -ne '127.0.0.1' -and $_.IPAddress -notlike '169.254.*'}} | Select-Object IPAddress,PrefixLength,PrefixOrigin); $gw=@(Get-NetRoute -InterfaceAlias {n} -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Where-Object {{$_.NextHop -and $_.NextHop -ne '0.0.0.0'}} | Sort-Object RouteMetric | Select-Object -ExpandProperty NextHop -Unique); $routes=@(Get-NetRoute -InterfaceAlias {n} -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {{$_.NextHop -and $_.NextHop -ne '0.0.0.0' -and $_.DestinationPrefix -ne '0.0.0.0/0' -and $_.Protocol -ne 'Dhcp'}} | Select-Object DestinationPrefix,NextHop); $dns=@(Get-DnsClientServerAddress -InterfaceAlias {n} -AddressFamily IPv4 -ErrorAction SilentlyContinue | ForEach-Object {{$_.ServerAddresses}}); $props=@(Get-NetAdapterAdvancedProperty -Name {n} -AllProperties -ErrorAction SilentlyContinue); $v=$props | Where-Object {{$_.RegistryKeyword -in @({','.join(ps_quote(k) for k in VLAN_KEYS)})}} | Select-Object -First 1; if(-not $v){{$v=$props | Where-Object {{$_.DisplayName -match '^VLAN\\s*ID$' -and $_.DisplayName -notmatch 'Priority'}} | Select-Object -First 1}}; $supports=$false; if($v -and ($v.RegistryKeyword -in @({','.join(ps_quote(k) for k in VLAN_KEYS)} ) -or $v.DisplayName -match '^VLAN\\s*ID$')){{$supports=$true}}; $vlanValue=''; if($v -and $v.RegistryValue -and [string]$v.RegistryValue[0] -ne ''){{$vlanValue=[string]$v.RegistryValue[0]; if($vlanValue -eq '0' -or ($v.DefaultRegistryValue -and $vlanValue -eq [string]$v.DefaultRegistryValue[0])){{$vlanValue=''}}}}; [pscustomobject]@{{Name=$a.Name;Guid=[string]$a.InterfaceGuid;Status=[string]$a.Status;LinkSpeed=[string]$a.LinkSpeed;Dhcp=[string]$ip.Dhcp;Addresses=$ad;Gateways=$gw;Routes=$routes;DnsServers=$dns;VlanKeyword=[string]$v.RegistryKeyword;VlanValue=$vlanValue;VlanSupported=$supports}} | ConvertTo-Json -Depth 6 -Compress"""
        result = json.loads(run_powershell(script) or "{}")
        if not isinstance(result.get("Addresses"), list):
            result["Addresses"] = [result["Addresses"]] if result.get("Addresses") else []
        for key in ("Gateways", "DnsServers", "Routes"):
            if not isinstance(result.get(key), list):
                result[key] = [result[key]] if result.get(key) else []
        parsed_routes = []
        for route in result["Routes"]:
            try:
                network = ipaddress.IPv4Network(route["DestinationPrefix"], strict=False)
                parsed_routes.append({"Destination": str(network.network_address),
                                      "Mask": str(network.netmask), "Gateway": route["NextHop"]})
            except (KeyError, ValueError):
                continue
        result["Routes"] = parsed_routes
        return result

    @staticmethod
    def snapshot_path(name: str) -> Path:
        slug = re.sub(r"[^A-Za-z0-9_-]+", "_", name).strip("_")[:40] or "adapter"
        digest = hashlib.sha256(name.encode("utf-8")).hexdigest()[:10]
        return SNAPSHOT_DIR / f"{slug}_{digest}.json"

    def ensure_snapshot(self, name: str, state: dict[str, Any] | None = None) -> None:
        path = self.snapshot_path(name)
        if path.exists():
            return
        SNAPSHOT_DIR.mkdir(exist_ok=True)
        # Import the legacy snapshot when it belongs to this adapter.
        if SNAPSHOT.exists():
            try:
                legacy = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
                if legacy.get("AdapterName") == name:
                    path.write_text(json.dumps(legacy, ensure_ascii=False, indent=2), encoding="utf-8")
                    return
            except (OSError, json.JSONDecodeError):
                pass
        state = state or self.state(name)
        path.write_text(json.dumps({**state, "AdapterName": name,
                                    "SavedAt": datetime.now().isoformat(timespec="seconds")},
                                   ensure_ascii=False, indent=2), encoding="utf-8")

    def apply(self, name: str, profile: dict[str, Any]) -> None:
        validate_profile(profile)
        current_state = self.state(name)
        changes_vlan = profile.get("VLAN") not in (None, "") or bool(current_state.get("VlanValue"))
        if changes_vlan and not is_administrator():
            raise PermissionError("Для изменения или сброса VLAN требуются права администратора. "
                                  "Закройте программу и запустите Start-Python.bat через «Запуск от имени администратора».")
        self.ensure_snapshot(name, current_state)
        p64 = base64.b64encode(json.dumps(profile, ensure_ascii=False).encode("utf-8")).decode("ascii")
        n = ps_quote(name)
        # VLAN property discovery uses known numeric ID properties only; the
        # Priority & VLAN toggle is deliberately never mistaken for a VLAN ID.
        keys = ",".join(ps_quote(k) for k in VLAN_KEYS)
        script = f"""$ErrorActionPreference='Stop'; $p=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{p64}')) | ConvertFrom-Json; $n={n}; $keys=@({keys}); $vlan=$null; if($null -ne $p.VLAN -and [string]$p.VLAN -ne ''){{$vlan=[int]$p.VLAN}}; $props=@(Get-NetAdapterAdvancedProperty -Name $n -AllProperties -ErrorAction Stop); $prop=$props | Where-Object {{$_.RegistryKeyword -in $keys}} | Select-Object -First 1; if(-not $prop){{$prop=$props | Where-Object {{$_.DisplayName -match '^VLAN\\s*ID$' -and $_.DisplayName -notmatch 'Priority'}} | Select-Object -First 1}}; if($null -ne $vlan){{if(-not $prop){{throw 'Драйвер не предоставляет подтверждённый параметр VLAN ID; IPv4 не менялся.'}}; Set-NetAdapterAdvancedProperty -Name $n -RegistryKeyword $prop.RegistryKeyword -RegistryValue ([string]$vlan) -ErrorAction Stop}} elseif($prop -and $prop.RegistryValue -and [string]$prop.RegistryValue[0] -ne ''){{if($prop.Optional){{Remove-NetAdapterAdvancedProperty -Name $n -RegistryKeyword $prop.RegistryKeyword -NoRestart -Confirm:$false -ErrorAction Stop}}elseif($prop.DefaultRegistryValue){{Set-NetAdapterAdvancedProperty -Name $n -RegistryKeyword $prop.RegistryKeyword -RegistryValue ([string]$prop.DefaultRegistryValue[0]) -NoRestart -ErrorAction Stop}}else{{throw 'Для найденного VLAN-драйвера не определён безопасный способ сброса; IPv4 не менялся.'}}}}; Start-Sleep -Seconds 2; $mode=[string]$p.Mode; if($mode -eq 'Network'){{& netsh.exe interface ipv4 set address ('name='+$n) source=static ('address='+[string]$p.IP) ('mask='+[string]$p.Mask) $(if($p.Gateway){{'gateway='+[string]$p.Gateway}}else{{'gateway=none'}}) store=persistent | Out-Null; if($LASTEXITCODE -ne 0){{throw 'netsh: ошибка настройки IPv4'}}}}; if($mode -in @('MultiAddress','Mixed')){{$a=@($p.Addresses); if($a.Count -gt 0){{& netsh.exe interface ipv4 set address ('name='+$n) source=static ('address='+[string]$a[0].IP) ('mask='+[string]$a[0].Mask) gateway=none store=active | Out-Null; if($LASTEXITCODE -ne 0){{throw 'netsh: ошибка IPv4'}}; for($i=1;$i -lt $a.Count;$i++){{& netsh.exe interface ipv4 add address ('name='+$n) ('address='+[string]$a[$i].IP) ('mask='+[string]$a[$i].Mask) store=active | Out-Null; if($LASTEXITCODE -ne 0){{throw 'netsh: ошибка дополнительного IPv4'}}}}}}}}; if($mode -in @('Routes','Mixed')){{foreach($r in @($p.Routes)){{$args=@('add',[string]$r.Destination,'mask',[string]$r.Mask,[string]$r.Gateway); if($p.Persistent){{$args=@('-p')+$args}}; & route.exe @args | Out-Null; if($LASTEXITCODE -ne 0){{throw ('Не удалось добавить маршрут '+$r.Destination)}}}}}}"""
        script += "; if($mode -eq 'DHCP'){Get-NetRoute -InterfaceAlias $n -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {$_.NextHop -and $_.NextHop -ne '0.0.0.0'} | Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue; Get-NetIPAddress -InterfaceAlias $n -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {$_.PrefixOrigin -ne 'WellKnown'} | Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue; Set-NetIPInterface -InterfaceAlias $n -AddressFamily IPv4 -Dhcp Enabled -ErrorAction Stop; if($p.DnsServers -and @($p.DnsServers).Count -gt 0){Set-DnsClientServerAddress -InterfaceAlias $n -ServerAddresses @($p.DnsServers) -ErrorAction Stop}else{Set-DnsClientServerAddress -InterfaceAlias $n -ResetServerAddresses -ErrorAction Stop}; Restart-NetAdapter -Name $n -Confirm:$false -ErrorAction Stop; Start-Sleep -Seconds 3; & ipconfig.exe /renew $n | Out-Null; if($p.Routes){foreach($r in $p.Routes){$ra=@('add',[string]$r.Destination,'mask',[string]$r.Mask,[string]$r.Gateway); if($p.Persistent){$ra=@('-p')+$ra}; & route.exe @ra | Out-Null; if($LASTEXITCODE -ne 0){throw ('Не удалось добавить маршрут '+$r.Destination)}}}}; if($mode -ne 'DHCP' -and $p.DnsServers -and @($p.DnsServers).Count -gt 0){Set-DnsClientServerAddress -InterfaceAlias $n -ServerAddresses @($p.DnsServers) -ErrorAction Stop}"
        script = script.replace(
            "elseif($prop -and $prop.RegistryValue -and [string]$prop.RegistryValue[0] -ne '')",
            "elseif($prop -and $prop.RegistryValue -and [string]$prop.RegistryValue[0] -ne '' -and [string]$prop.RegistryValue[0] -ne '0' -and (-not $prop.DefaultRegistryValue -or [string]$prop.RegistryValue[0] -ne [string]$prop.DefaultRegistryValue[0]))",
        )
        run_powershell(script, timeout=90)

    def dhcp(self, name: str) -> None:
        self.ensure_snapshot(name)
        n = ps_quote(name)
        keys = ",".join(ps_quote(k) for k in VLAN_KEYS)
        script = f"""$ErrorActionPreference='Stop'; $n={n}; $v=Get-NetAdapterAdvancedProperty -Name $n -AllProperties -ErrorAction Stop | Where-Object {{$_.RegistryKeyword -in @({keys})}} | Select-Object -First 1; if($v -and $v.RegistryValue -and [string]$v.RegistryValue[0] -ne ''){{if($v.Optional){{Remove-NetAdapterAdvancedProperty -Name $n -RegistryKeyword $v.RegistryKeyword -NoRestart -Confirm:$false -ErrorAction Stop}}elseif($v.DefaultRegistryValue){{Set-NetAdapterAdvancedProperty -Name $n -RegistryKeyword $v.RegistryKeyword -RegistryValue ([string]$v.DefaultRegistryValue[0]) -NoRestart -ErrorAction Stop}}else{{throw 'Способ сброса VLAN не подтверждён; DHCP не включён.'}}}}; Get-NetRoute -InterfaceAlias $n -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue; Get-NetIPAddress -InterfaceAlias $n -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {{$_.PrefixOrigin -ne 'WellKnown'}} | Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue; Set-NetIPInterface -InterfaceAlias $n -AddressFamily IPv4 -Dhcp Enabled -ErrorAction Stop; Set-DnsClientServerAddress -InterfaceAlias $n -ResetServerAddresses -ErrorAction Stop; Restart-NetAdapter -Name $n -Confirm:$false -ErrorAction Stop; Start-Sleep -Seconds 3; & ipconfig.exe /renew $n | Out-Null"""
        run_powershell(script, timeout=35)

    def restore(self, name: str) -> None:
        path = self.snapshot_path(name)
        if not path.exists() and SNAPSHOT.exists():
            try:
                legacy = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
                if legacy.get("AdapterName") == name:
                    path = SNAPSHOT
            except (OSError, json.JSONDecodeError):
                pass
        if not path.exists():
            raise RuntimeError("Снимок исходного состояния ещё не создан")
        snap = json.loads(path.read_text(encoding="utf-8"))
        if snap.get("AdapterName") != name:
            raise RuntimeError(f"Снимок относится к адаптеру {snap.get('AdapterName')}")
        s64 = base64.b64encode(json.dumps(snap).encode()).decode()
        n = ps_quote(name)
        script = f"""$ErrorActionPreference='Stop'; function Mask([int]$p){{$b=('1'*$p).PadRight(32,'0'); return (0,8,16,24 | ForEach-Object {{[Convert]::ToInt32($b.Substring($_,8),2)}}) -join '.'}}; $s=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{s64}')) | ConvertFrom-Json; $n={n}; $v=Get-NetAdapterAdvancedProperty -Name $n -AllProperties -ErrorAction Stop | Where-Object {{$_.RegistryKeyword -in @({','.join(ps_quote(k) for k in VLAN_KEYS)})}} | Select-Object -First 1; if($v -and $v.RegistryValue -and [string]$v.RegistryValue[0] -ne ''){{if($v.Optional){{Remove-NetAdapterAdvancedProperty -Name $n -RegistryKeyword $v.RegistryKeyword -NoRestart -Confirm:$false -ErrorAction Stop}}elseif($v.DefaultRegistryValue){{Set-NetAdapterAdvancedProperty -Name $n -RegistryKeyword $v.RegistryKeyword -RegistryValue ([string]$v.DefaultRegistryValue[0]) -NoRestart -ErrorAction Stop}}else{{throw 'Нельзя безопасно отключить текущий VLAN; восстановление остановлено.'}}}}; Get-NetRoute -InterfaceAlias $n -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {{$_.NextHop -and $_.NextHop -ne '0.0.0.0'}} | Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue; Get-NetIPAddress -InterfaceAlias $n -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {{$_.PrefixOrigin -ne 'WellKnown'}} | Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue; if($s.VlanKeyword -and $s.VlanValue){{Set-NetAdapterAdvancedProperty -Name $n -RegistryKeyword $s.VlanKeyword -RegistryValue ([string]$s.VlanValue) -NoRestart -ErrorAction Stop}}; if($s.Dhcp -eq 'Enabled'){{Set-NetIPInterface -InterfaceAlias $n -AddressFamily IPv4 -Dhcp Enabled -ErrorAction Stop}}else{{$a=@($s.Addresses); if($a.Count -gt 0){{$mask=Mask ([int]$a[0].PrefixLength); $gw=@($s.Gateways) | Select-Object -First 1; & netsh.exe interface ipv4 set address ('name='+$n) source=static ('address='+$a[0].IPAddress) ('mask='+$mask) $(if($gw){{'gateway='+$gw}}else{{'gateway=none'}}) store=persistent | Out-Null; for($i=1;$i -lt $a.Count;$i++){{$mask=Mask ([int]$a[$i].PrefixLength); & netsh.exe interface ipv4 add address ('name='+$n) ('address='+$a[$i].IPAddress) ('mask='+$mask) store=persistent | Out-Null}}}}}}; foreach($r in @($s.Routes)){{& route.exe add ([string]$r.Destination) mask ([string]$r.Mask) ([string]$r.Gateway) | Out-Null}}; if(@($s.DnsServers).Count){{Set-DnsClientServerAddress -InterfaceAlias $n -ServerAddresses @($s.DnsServers) -ErrorAction Stop}}else{{Set-DnsClientServerAddress -InterfaceAlias $n -ResetServerAddresses -ErrorAction Stop}}; Restart-NetAdapter -Name $n -Confirm:$false -ErrorAction Stop"""
        run_powershell(script, timeout=60)


class ProfileEditor(tk.Toplevel):
    def __init__(self, master: "NetworkManagerApp", profile: dict[str, Any] | None = None,
                 creating: bool = False):
        super().__init__(master)
        self.title("Создание профиля" if creating or profile is None else "Изменение профиля")
        self.geometry("700x680")
        self.minsize(620, 480)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.transient(master)
        self.grab_set()
        self.result: dict[str, Any] | None = None
        self.initial = dict(profile or {})
        self.name_var = tk.StringVar(value=self.initial.get("Name", ""))
        self.category_var = tk.StringVar(value=self.initial.get("Category", "Сетевые профили"))
        initial_mode = self.initial.get("Mode", "Network")
        self.mode_var = tk.StringVar(value=MODE_LABELS.get(initial_mode, MODE_LABELS["Network"]))
        vlan = self.initial.get("VLAN")
        self.vlan_var = tk.StringVar(value="" if vlan is None else str(vlan))
        self.ip_var = tk.StringVar(value=self.initial.get("IP", ""))
        self.mask_var = tk.StringVar(value=self.initial.get("Mask", "255.255.255.0"))
        self.gateway_var = tk.StringVar(value=self.initial.get("Gateway", ""))
        self.device_var = tk.StringVar(value=self.initial.get("DeviceName", ""))
        self.ping_var = tk.StringVar(value=self.initial.get("PingTarget", ""))
        self.dns_var = tk.StringVar(value=", ".join(self.initial.get("DnsServers", [])))
        self.persistent_var = tk.BooleanVar(value=bool(self.initial.get("Persistent", False)))
        self.address_rows: list[tuple[tk.StringVar, tk.StringVar]] = []
        self.route_rows: list[tuple[tk.StringVar, tk.StringVar, tk.StringVar]] = []
        self.address_cache = [(x.get("IP", ""), x.get("Mask", "255.255.255.0")) for x in self.initial.get("Addresses", [])]
        self.route_cache = [(x.get("Destination", ""), x.get("Mask", "255.255.255.0"), x.get("Gateway", "")) for x in self.initial.get("Routes", [])]

        header = ttk.Frame(self, padding=(20, 12, 20, 6)); header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Настройте профиль", font=("Segoe UI", 17, "bold")).pack(anchor="w", pady=(2, 2))
        ttk.Label(header, text="Поля с * обязательны. Маску можно указать как 255.255.255.0 или /24.",
                  foreground="#687386").pack(anchor="w", padx=20, pady=(0, 12))

        body = ttk.Frame(self); body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1); body.rowconfigure(0, weight=1)
        canvas = tk.Canvas(body, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew"); scrollbar.grid(row=0, column=1, sticky="ns")
        self.form = ttk.Frame(canvas, padding=(18, 0, 18, 6))
        form_window = canvas.create_window((0, 0), window=self.form, anchor="nw")
        self.form.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(form_window, width=e.width))
        self.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        main = ttk.LabelFrame(self.form, text="Основное", padding=12); main.pack(fill="x", pady=5)
        main.columnconfigure(1, weight=1); main.columnconfigure(3, weight=1)
        self.add_labeled_entry(main, "Название *", self.name_var, 0, 0)
        self.add_labeled_entry(main, "Категория", self.category_var, 0, 2)
        ttk.Label(main, text="Что настроить? *").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=7)
        mode = ttk.Combobox(main, textvariable=self.mode_var, state="readonly",
                            values=tuple(MODE_LABELS.values()), width=25)
        mode.grid(row=1, column=1, sticky="ew", padx=(0, 14), pady=7)
        mode.bind("<<ComboboxSelected>>", lambda _e: self.build_mode_fields())
        self.add_labeled_entry(main, "VLAN ID", self.vlan_var, 1, 2)
        ttk.Label(main, text="Оставьте VLAN пустым, если тег VLAN не нужен.", foreground="#687386").grid(
            row=2, column=1, columnspan=3, sticky="w", pady=(0, 2))
        self.mode_help = ttk.Label(main, text="", wraplength=620, justify="left", foreground="#526174")
        self.mode_help.grid(row=3, column=0, columnspan=4, sticky="w", pady=(8, 2))

        self.mode_fields = ttk.Frame(self.form); self.mode_fields.pack(fill="x", pady=3)
        self.build_mode_fields()
        extras = ttk.LabelFrame(self.form, text="Дополнительно", padding=12); extras.pack(fill="x", pady=5)
        extras.columnconfigure(1, weight=1); extras.columnconfigure(3, weight=1)
        self.add_labeled_entry(extras, "Устройство", self.device_var, 0, 0)
        self.add_labeled_entry(extras, "Проверить связь с", self.ping_var, 0, 2)
        ttk.Label(extras, text="DNS-серверы").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=7)
        ttk.Entry(extras, textvariable=self.dns_var).grid(row=1, column=1, columnspan=3, sticky="ew", pady=7)
        ttk.Label(extras, text="PingTarget проверяется после применения. DNS можно перечислить через запятую.",
                  foreground="#687386").grid(row=2, column=1, columnspan=3, sticky="w", pady=(2, 0))
        footer = ttk.Frame(self, padding=(18, 8, 18, 10)); footer.grid(row=2, column=0, sticky="ew")
        ttk.Label(footer, text="Прокрутите форму, чтобы увидеть все поля.", foreground="#687386").pack(side="left")
        ttk.Button(footer, text="Отмена", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(footer, text="Сохранить профиль", command=self.save).pack(side="right")

    @staticmethod
    def add_labeled_entry(parent: ttk.Frame, label: str, variable: tk.StringVar,
                          row: int, column: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=7)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=column + 1, sticky="ew", padx=(0, 14), pady=7)

    def build_mode_fields(self) -> None:
        if hasattr(self, "address_rows"):
            self.address_cache = [(ip.get(), mask.get()) for ip, mask in self.address_rows]
            self.route_cache = [(d.get(), mask.get(), gateway.get()) for d, mask, gateway in self.route_rows]
        for child in self.mode_fields.winfo_children(): child.destroy()
        self.address_rows.clear(); self.route_rows.clear()
        mode = mode_from_label(self.mode_var.get())
        self.mode_help.configure(text=MODE_DESCRIPTIONS.get(mode, ""))
        if mode == "Network":
            box = ttk.LabelFrame(self.mode_fields, text="Настройки IPv4", padding=12); box.pack(fill="x")
            box.columnconfigure(1, weight=1); box.columnconfigure(3, weight=1)
            self.add_labeled_entry(box, "IPv4 адрес *", self.ip_var, 0, 0)
            self.add_labeled_entry(box, "Маска *", self.mask_var, 0, 2)
            self.add_labeled_entry(box, "Шлюз", self.gateway_var, 1, 0)
            ttk.Label(box, text="Например: 192.168.1.20 / 255.255.255.0", foreground="#687386").grid(
                row=2, column=1, columnspan=3, sticky="w")
        if mode == "DHCP":
            box = ttk.LabelFrame(self.mode_fields, text="Автоматическая настройка", padding=12); box.pack(fill="x", pady=4)
            ttk.Label(box, text="IP-адрес, шлюз и DNS будут получены от DHCP-сервера.",
                      wraplength=600, justify="left").pack(anchor="w")
        if mode in ("MultiAddress", "Mixed"):
            box = ttk.LabelFrame(self.mode_fields, text="Дополнительные IPv4 адреса", padding=10); box.pack(fill="x", pady=4)
            self.address_list = ttk.Frame(box); self.address_list.pack(fill="x")
            for ip, mask in self.address_cache: self.add_address_row(ip, mask)
            ttk.Button(box, text="＋ Добавить IP-адрес", command=self.add_address_row).pack(anchor="w", pady=(8, 0))
            if not self.address_rows: self.add_address_row()
        if mode in ("Routes", "Mixed", "DHCP"):
            box = ttk.LabelFrame(self.mode_fields, text="Маршруты", padding=10); box.pack(fill="x", pady=4)
            self.route_list = ttk.Frame(box); self.route_list.pack(fill="x")
            for destination, mask, gateway in self.route_cache:
                self.add_route_row(destination, mask, gateway)
            ttk.Button(box, text="＋ Добавить маршрут", command=self.add_route_row).pack(anchor="w", pady=(8, 0))
            if mode != "DHCP" and not self.route_rows: self.add_route_row()
            ttk.Checkbutton(box, text="Сделать маршруты постоянными", variable=self.persistent_var).pack(anchor="w", pady=(8, 0))

    def add_address_row(self, ip: str = "", mask: str = "255.255.255.0") -> None:
        row = ttk.Frame(self.address_list); row.pack(fill="x", pady=2)
        ip_var, mask_var = tk.StringVar(value=ip), tk.StringVar(value=mask)
        ttk.Label(row, text="IP").pack(side="left", padx=(0, 5)); ttk.Entry(row, textvariable=ip_var, width=23).pack(side="left", padx=(0, 12))
        ttk.Label(row, text="Маска").pack(side="left", padx=(0, 5)); ttk.Entry(row, textvariable=mask_var, width=23).pack(side="left", padx=(0, 10))
        record = (ip_var, mask_var); self.address_rows.append(record)
        ttk.Button(row, text="Удалить", command=lambda: self.remove_row(row, record, self.address_rows)).pack(side="left")

    def add_route_row(self, destination: str = "", mask: str = "255.255.255.0", gateway: str = "") -> None:
        row = ttk.Frame(self.route_list); row.pack(fill="x", pady=2)
        d_var, m_var, g_var = tk.StringVar(value=destination), tk.StringVar(value=mask), tk.StringVar(value=gateway)
        for label, var, width in (("Сеть", d_var, 18), ("Маска", m_var, 18), ("Шлюз", g_var, 18)):
            ttk.Label(row, text=label).pack(side="left", padx=(0, 4)); ttk.Entry(row, textvariable=var, width=width).pack(side="left", padx=(0, 8))
        record = (d_var, m_var, g_var); self.route_rows.append(record)
        ttk.Button(row, text="Удалить", command=lambda: self.remove_row(row, record, self.route_rows)).pack(side="left")

    @staticmethod
    def remove_row(widget: ttk.Frame, record: tuple, collection: list) -> None:
        if record in collection: collection.remove(record)
        widget.destroy()

    def save(self) -> None:
        try:
            p = dict(self.initial)
            p.update({"Name": self.name_var.get().strip(), "Category": self.category_var.get().strip(),
                      "Mode": self.mode_var.get(), "VLAN": int(self.vlan_var.get()) if self.vlan_var.get().strip() else None,
                      "DeviceName": self.device_var.get().strip(), "PingTarget": self.ping_var.get().strip()})
            if not p["DeviceName"]: p.pop("DeviceName")
            if not p["PingTarget"]: p.pop("PingTarget")
            dns_servers = [x.strip() for x in re.split(r"[,;\s]+", self.dns_var.get()) if x.strip()]
            if dns_servers: p["DnsServers"] = dns_servers
            else: p.pop("DnsServers", None)
            mode = mode_from_label(self.mode_var.get())
            p["Mode"] = mode
            if mode == "Network":
                p.update(IP=self.ip_var.get().strip(), Mask=self.mask_var.get().strip(), Gateway=self.gateway_var.get().strip())
                p.pop("Addresses", None); p.pop("Routes", None); p.pop("Persistent", None)
            elif mode in ("MultiAddress", "Mixed"):
                p["Addresses"] = [{"IP": ip.get().strip(), "Mask": mask.get().strip()} for ip, mask in self.address_rows]
                p.pop("IP", None); p.pop("Mask", None); p.pop("Gateway", None)
                if mode == "MultiAddress": p.pop("Routes", None); p.pop("Persistent", None)
            elif mode == "DHCP":
                p.pop("IP", None); p.pop("Mask", None); p.pop("Gateway", None); p.pop("Addresses", None)
                if self.route_rows:
                    p["Routes"] = [{"Destination": d.get().strip(), "Mask": mask.get().strip(), "Gateway": gateway.get().strip()}
                                    for d, mask, gateway in self.route_rows]
                    p["Persistent"] = self.persistent_var.get()
                else:
                    p.pop("Routes", None); p.pop("Persistent", None)
            else:
                p.pop("IP", None); p.pop("Mask", None); p.pop("Gateway", None); p.pop("Addresses", None)
            if mode in ("Routes", "Mixed"):
                p["Routes"] = [{"Destination": d.get().strip(), "Mask": mask.get().strip(), "Gateway": gateway.get().strip()}
                                for d, mask, gateway in self.route_rows]
                p["Persistent"] = self.persistent_var.get()
            validate_profile(p)
            self.result = p
            self.destroy()
        except Exception as e:
            messagebox.showerror("Ошибка профиля", str(e), parent=self)


class NetworkManagerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} · Python GUI")
        self.geometry("1220x780")
        self.minsize(900, 600)
        self.backend = NetworkBackend()
        self.adapters: list[dict[str, Any]] = []
        self.object_path: Path | None = None
        self.profiles: list[dict[str, Any]] = []
        self.settings = self.load_settings()
        self.adapter_var = tk.StringVar()
        self.object_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Загрузка адаптеров…")
        self.build_ui()
        self.refresh_adapters()

    @staticmethod
    def load_settings() -> dict[str, Any]:
        try:
            return json.loads(SETTINGS.read_text(encoding="utf-8"))
        except Exception:
            return {"dark": True, "show_device": True, "enable_ping": True}

    def save_settings(self) -> None:
        SETTINGS.write_text(json.dumps(self.settings, ensure_ascii=False, indent=2), encoding="utf-8")

    def build_ui(self) -> None:
        self.configure(padx=14, pady=12)
        head = ttk.Frame(self); head.pack(fill="x", pady=(0, 10))
        ttk.Label(head, text="Network Service Manager", font=("Segoe UI", 20, "bold")).pack(side="left")
        ttk.Label(head, text="Управление сетевыми профилями", font=("Segoe UI", 10)).pack(side="left", padx=16, pady=(9, 0))
        ttk.Button(head, text="⚙ Настройки", command=self.app_settings).pack(side="right", padx=(8, 0))
        ttk.Button(head, text="Светлая / тёмная тема", command=self.toggle_theme).pack(side="right")
        ttk.Label(self, text="Выберите адаптер и объект, затем дважды щёлкните профиль или нажмите «Применить профиль».",
                  foreground="#687386").pack(anchor="w", pady=(0, 7))
        adapter_bar = ttk.LabelFrame(self, text="ШАГ 1 · Сетевой адаптер", padding=10); adapter_bar.pack(fill="x", pady=5)
        self.adapter_combo = ttk.Combobox(adapter_bar, textvariable=self.adapter_var, state="readonly", width=44)
        self.adapter_combo.pack(side="left", padx=(0, 8)); self.adapter_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_state())
        ttk.Button(adapter_bar, text="Обновить", command=self.refresh_adapters).pack(side="left")
        self.state_label = ttk.Label(adapter_bar, text=""); self.state_label.pack(side="left", padx=18)
        recovery = ttk.LabelFrame(self, text="Восстановление сети", padding=(8, 6)); recovery.pack(fill="x", pady=5)
        ttk.Button(recovery, text="Включить DHCP и DNS", style="Accent.TButton",
                   command=self.restore_dhcp).pack(side="left", padx=(0, 8))
        ttk.Button(recovery, text="Вернуть исходные настройки", command=self.restore_original).pack(side="left")
        ttk.Label(recovery, text="Исходные параметры сохраняются отдельно для каждого адаптера.",
                  foreground="#687386").pack(side="left", padx=12)
        objbar = ttk.LabelFrame(self, text="ШАГ 2 · Объект с профилями", padding=10); objbar.pack(fill="x", pady=5)
        self.object_combo = ttk.Combobox(objbar, textvariable=self.object_var, state="readonly", width=40)
        self.object_combo.pack(side="left", padx=(0, 8)); self.object_combo.bind("<<ComboboxSelected>>", lambda _e: self.load_object())
        ttk.Button(objbar, text="Новый объект", command=self.new_object).pack(side="left", padx=3)
        ttk.Button(objbar, text="Переименовать", command=self.rename_object).pack(side="left", padx=3)
        ttk.Button(objbar, text="Сохранить текущие настройки как профиль",
                   command=self.save_current_as_profile).pack(side="left", padx=(10, 3))
        self.pane = ttk.Panedwindow(self, orient="horizontal"); self.pane.pack(fill="both", expand=True, pady=8)
        left = ttk.Frame(self.pane); right = ttk.Frame(self.pane); self.pane.add(left, weight=3); self.pane.add(right, weight=2)
        ttk.Label(left, text="ШАГ 3 · Выберите профиль", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 5))
        cols = ("name", "mode", "vlan", "details")
        table = ttk.Frame(left); table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table, columns=cols, show="headings", selectmode="browse")
        for c, label, width in (("name", "Название", 180), ("mode", "Тип профиля", 155), ("vlan", "VLAN", 65), ("details", "Параметры", 270)):
            self.tree.heading(c, text=label); self.tree.column(c, width=width, minwidth=70, anchor="w", stretch=(c == "details"))
        vertical = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        horizontal = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.tree.grid(row=0, column=0, sticky="nsew"); vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        table.rowconfigure(0, weight=1); table.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self.show_profile())
        self.tree.bind("<Double-1>", lambda _e: self.apply_selected())
        actions = ttk.Frame(left); actions.pack(fill="x", pady=8)
        ttk.Button(actions, text="▶  Применить профиль", style="Accent.TButton", command=self.apply_selected).pack(side="left", padx=(0, 10))
        ttk.Button(actions, text="＋ Создать профиль", command=self.add_profile).pack(side="left", padx=(0, 5))
        ttk.Button(actions, text="Изменить", command=self.edit_profile).pack(side="left", padx=(0, 5))
        ttk.Button(actions, text="Удалить", command=self.delete_profile).pack(side="left")
        ttk.Label(right, text="Состояние адаптера", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 5))
        self.state_text = tk.Text(right, height=11, wrap="word", state="disabled", font=("Consolas", 10))
        self.state_text.pack(fill="x", pady=(0, 10))
        ttk.Label(right, text="Выбранный профиль", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 5))
        self.detail_text = tk.Text(right, wrap="word", state="disabled", font=("Consolas", 10))
        self.detail_text.pack(fill="both", expand=True)
        bottom = ttk.Frame(self); bottom.pack(fill="x")
        ttk.Button(bottom, text="Проверить связь (ping)", command=self.manual_ping).pack(side="left", padx=(0, 6))
        ttk.Label(bottom, textvariable=self.status_var, anchor="e").pack(side="right", fill="x", expand=True)
        if self.settings.get("dark"):
            self.set_dark()
        else:
            self.set_light()

    def set_dark(self) -> None:
        style = ttk.Style(self); style.theme_use("clam")
        bg, panel, fg, select = "#171b22", "#222832", "#e8edf4", "#235b8e"
        self.configure(bg=bg)
        style.configure(".", background=bg, foreground=fg, fieldbackground=panel, bordercolor="#3c4654")
        style.configure("TFrame", background=bg); style.configure("TLabel", background=bg, foreground=fg)
        style.configure("TLabelframe", background=bg, foreground=fg); style.configure("TLabelframe.Label", background=bg, foreground=fg)
        style.configure("TButton", background=panel, foreground=fg, padding=6); style.map("TButton", background=[("active", select)])
        style.configure("Accent.TButton", background="#1769aa", foreground="white", padding=(10, 7), font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#0e548c")])
        style.configure("TCombobox", fieldbackground=panel, foreground=fg, arrowcolor=fg)
        style.configure("Treeview", background=panel, foreground=fg, fieldbackground=panel, rowheight=27)
        style.map("Treeview", background=[("selected", select)], foreground=[("selected", "white")])
        for widget in (getattr(self, "state_text", None), getattr(self, "detail_text", None)):
            if widget: widget.configure(bg=panel, fg=fg, insertbackground=fg)

    def toggle_theme(self) -> None:
        self.settings["dark"] = not self.settings.get("dark", True); self.save_settings()
        if self.settings["dark"]: self.set_dark()
        else: self.set_light()

    def set_light(self) -> None:
        style = ttk.Style(self)
        style.theme_use("vista" if "vista" in style.theme_names() else "clam")
        self.configure(bg="#f0f0f0")
        for widget in (getattr(self, "state_text", None), getattr(self, "detail_text", None)):
            if widget: widget.configure(bg="white", fg="#20242b", insertbackground="#20242b")

    def selected_adapter(self) -> str:
        name = self.adapter_var.get().split("  ·  ")[0]
        if not name: raise RuntimeError("Сначала выберите сетевой адаптер")
        return name

    def refresh_adapters(self) -> None:
        def work():
            try:
                items = self.backend.adapters()
                self.after(0, lambda: self.show_adapters(items))
            except Exception as e:
                error_text = str(e)
                self.after(0, lambda: (self.status_var.set("Ошибка чтения адаптеров"), messagebox.showerror("Адаптеры", f"{error_text}\n\nЗапустите приложение в Windows.")))
        threading.Thread(target=work, daemon=True).start()

    def show_adapters(self, items: list[dict[str, Any]]) -> None:
        self.adapters = items
        values = [f"{a['Name']}  ·  {a.get('Status','')}  ·  {a.get('InterfaceDescription','')}" for a in items]
        self.adapter_combo["values"] = values
        if values and (not self.adapter_var.get() or not any(v.startswith(self.adapter_var.get().split("  ·  ")[0] + "  ·") for v in values)):
            self.adapter_var.set(values[0])
        self.status_var.set(f"Адаптеров: {len(items)}")
        self.refresh_state()
        self.load_objects()

    def refresh_state(self) -> None:
        if not self.adapter_var.get(): return
        name = self.selected_adapter()
        self.status_var.set("Получение состояния…")
        def work():
            try:
                s = self.backend.state(name)
                try:
                    self.backend.ensure_snapshot(name, s)
                    snapshot_error = ""
                except Exception as e:
                    snapshot_error = str(e)
                    logging.exception("Не удалось сохранить исходные параметры адаптера %s", name)
                def show():
                    ips = ", ".join(f"{x['IPAddress']}/{x['PrefixLength']}" for x in s.get("Addresses", [])) or "—"
                    self.state_label.configure(text=f"{s.get('Status')} · VLAN {s.get('VlanValue') or 'нет'} · {ips}")
                    rows = [f"Адаптер: {s.get('Name')}", f"Состояние: {s.get('Status')}", f"Скорость: {s.get('LinkSpeed')}",
                            f"VLAN: {s.get('VlanValue') or 'отсутствует'} ({s.get('VlanKeyword') or 'не обнаружен'})",
                            f"VLAN ID подтверждён: {'да' if s.get('VlanSupported') else 'нет'}", f"DHCP: {s.get('Dhcp') or 'неизвестно'}",
                            f"IPv4: {ips}", f"Шлюз: {', '.join(s.get('Gateways', [])) or '—'}", f"DNS: {', '.join(s.get('DnsServers', [])) or '—'}",
                            f"Маршрутов через шлюз: {len(s.get('Routes', []))}"]
                    self.set_text(self.state_text, "\n".join(rows))
                    self.status_var.set(f"Не сохранён снимок исходных настроек: {snapshot_error}" if snapshot_error else "Состояние обновлено")
                self.after(0, show)
            except Exception as e:
                error_text = str(e)
                self.after(0, lambda: self.status_var.set(f"Не удалось прочитать состояние: {error_text}"))
        threading.Thread(target=work, daemon=True).start()

    @staticmethod
    def set_text(widget: tk.Text, value: str) -> None:
        widget.configure(state="normal"); widget.delete("1.0", "end"); widget.insert("1.0", value); widget.configure(state="disabled")

    def load_objects(self) -> None:
        PROFILES_DIR.mkdir(exist_ok=True)
        names = sorted(p.stem for p in PROFILES_DIR.glob("*.json") if not p.name.endswith(".json.bak"))
        self.object_combo["values"] = names
        if names:
            if self.object_var.get() not in names: self.object_var.set(names[0])
            self.load_object()
        else:
            self.object_var.set(""); self.object_path = None; self.profiles = []; self.render_profiles()

    def load_object(self) -> None:
        if not self.object_var.get(): return
        path = PROFILES_DIR / f"{self.object_var.get()}.json"
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
            self.profiles = data if isinstance(data, list) else [data]
            self.object_path = path; self.render_profiles()
        except Exception as e:
            messagebox.showerror("Ошибка объекта", str(e)); self.profiles = []; self.render_profiles()

    def render_profiles(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for i, p in enumerate(self.profiles):
            mode = p.get("Mode") or ("Mixed" if "Addresses" in p and "Routes" in p else "MultiAddress" if "Addresses" in p else "Routes" if "Routes" in p else "Network")
            mode_label = MODE_LABELS.get(mode, mode)
            vlan = f"{p.get('VLAN')}" if p.get("VLAN") else "нет"
            if mode == "Network": detail = f"{p.get('IP','')} / {p.get('Mask','')}  GW {p.get('Gateway') or '—'}"
            elif mode == "DHCP": detail = f"Автоматический IP · {len(p.get('Routes', []))} маршрутов"
            else: detail = f"{len(p.get('Addresses', []))} IPv4 · {len(p.get('Routes', []))} маршрутов"
            if p.get("Category"): detail = f"[{p['Category']}] {detail}"
            self.tree.insert("", "end", iid=str(i), values=(p.get("Name", "Без названия"), mode_label, vlan, detail))
        self.set_text(self.detail_text, "Выберите профиль для просмотра параметров.")

    def current_profile(self) -> tuple[int, dict[str, Any]] | None:
        selected = self.tree.selection()
        if not selected: return None
        i = int(selected[0]); return i, self.profiles[i]

    def show_profile(self) -> None:
        item = self.current_profile()
        if item:
            _, p = item
            shown = dict(p)
            if not self.settings.get("show_device", True): shown.pop("DeviceName", None)
            self.set_text(self.detail_text, json.dumps(shown, ensure_ascii=False, indent=2))

    def save_profiles(self) -> None:
        if not self.object_path: raise RuntimeError("Создайте или выберите объект")
        if self.object_path.exists():
            self.object_path.with_suffix(self.object_path.suffix + ".bak").write_bytes(self.object_path.read_bytes())
        self.object_path.write_text(json.dumps(self.profiles, ensure_ascii=False, indent=2), encoding="utf-8")
        self.render_profiles()

    def new_object(self) -> None:
        name = simpledialog.askstring("Новый объект", "Название объекта:", parent=self)
        if name is None: return
        name = name.strip()
        if not name or name in (".", "..") or re.search(r'[<>:"/\\|?*]', name):
            messagebox.showerror("Объект", "Недопустимое имя объекта."); return
        path = PROFILES_DIR / f"{name}.json"
        if path.exists(): messagebox.showerror("Объект", "Объект с таким именем уже есть."); return
        path.write_text("[]", encoding="utf-8"); self.object_var.set(name); self.load_objects()

    def rename_object(self) -> None:
        if not self.object_path: return
        name = simpledialog.askstring("Переименовать объект", "Новое имя:", initialvalue=self.object_path.stem, parent=self)
        if not name: return
        if re.search(r'[<>:"/\\|?*]', name) or (PROFILES_DIR / f"{name}.json").exists():
            messagebox.showerror("Объект", "Недопустимое или занятое имя."); return
        old = self.object_path; new = PROFILES_DIR / f"{name}.json"; old.rename(new)
        self.object_var.set(name); self.load_objects()

    def add_profile(self) -> None:
        if not self.object_path:
            messagebox.showinfo("Объект", "Сначала создайте объект."); return
        editor = ProfileEditor(self, creating=True); self.wait_window(editor)
        if editor.result:
            self.profiles.append(editor.result)
            try: self.save_profiles()
            except Exception as e: messagebox.showerror("Сохранение", str(e))

    def edit_profile(self) -> None:
        item = self.current_profile()
        if not item: return
        i, p = item; editor = ProfileEditor(self, p); self.wait_window(editor)
        if editor.result:
            self.profiles[i] = editor.result
            try: self.save_profiles()
            except Exception as e: messagebox.showerror("Сохранение", str(e))

    def save_current_as_profile(self) -> None:
        if not self.object_path:
            messagebox.showinfo("Сохранить профиль", "Сначала выберите или создайте объект для профиля.", parent=self)
            return
        try:
            name = self.selected_adapter()
        except RuntimeError as e:
            messagebox.showerror("Сохранить профиль", str(e), parent=self)
            return
        self.status_var.set("Чтение текущих сетевых настроек…")
        def work():
            try:
                state = self.backend.state(name)
                self.after(0, lambda: self.open_current_profile_editor(name, state))
            except Exception as e:
                error_text = str(e)
                self.after(0, lambda: (self.status_var.set("Не удалось прочитать настройки"),
                                       messagebox.showerror("Сохранить профиль", error_text, parent=self)))
        threading.Thread(target=work, daemon=True).start()

    def open_current_profile_editor(self, adapter_name: str, state: dict[str, Any]) -> None:
        addresses = state.get("Addresses", [])
        routes = list(state.get("Routes", []))
        gateways = state.get("Gateways", [])
        gateway = gateways[0] if gateways else ""
        profile: dict[str, Any] = {
            "Name": f"Текущие настройки {adapter_name}",
            "Category": "Сохранённые настройки",
            "VLAN": int(state["VlanValue"]) if str(state.get("VlanValue", "")).isdigit() else None,
        }
        dhcp_enabled = state.get("Dhcp") == "Enabled"
        if dhcp_enabled:
            profile["Mode"] = "DHCP"
            if routes:
                profile.update({"Routes": routes, "Persistent": False})
        elif gateway and (not addresses or len(addresses) > 1 or routes):
            routes.insert(0, {"Destination": "0.0.0.0", "Mask": "0.0.0.0", "Gateway": gateway})
        if dhcp_enabled:
            pass
        elif len(addresses) == 1 and not routes:
            address = addresses[0]
            profile.update({"Mode": "Network", "IP": address["IPAddress"],
                            "Mask": mask_from(str(address["PrefixLength"])), "Gateway": gateway})
        elif addresses and routes:
            profile.update({"Mode": "Mixed",
                            "Addresses": [{"IP": a["IPAddress"], "Mask": mask_from(str(a["PrefixLength"]))} for a in addresses],
                            "Routes": routes, "Persistent": False})
        elif len(addresses) > 1:
            profile.update({"Mode": "MultiAddress",
                            "Addresses": [{"IP": a["IPAddress"], "Mask": mask_from(str(a["PrefixLength"]))} for a in addresses]})
        elif routes:
            profile.update({"Mode": "Routes", "Routes": routes, "Persistent": False})
        else:
            messagebox.showwarning("Сохранить профиль", "На адаптере не найдено IPv4-адресов или маршрутов, которые можно сохранить.", parent=self)
            self.status_var.set("Нет параметров для сохранения")
            return
        if not dhcp_enabled and state.get("DnsServers"):
            profile["DnsServers"] = state["DnsServers"]
        editor = ProfileEditor(self, profile, creating=True)
        self.wait_window(editor)
        if editor.result:
            self.profiles.append(editor.result)
            try:
                self.save_profiles()
                self.status_var.set("Текущие настройки сохранены как новый профиль")
            except Exception as e:
                messagebox.showerror("Сохранение", str(e), parent=self)

    def delete_profile(self) -> None:
        item = self.current_profile()
        if not item: return
        i, p = item
        if messagebox.askyesno("Удалить профиль", f"Удалить профиль «{p.get('Name')}»?", parent=self):
            del self.profiles[i]
            try: self.save_profiles()
            except Exception as e: messagebox.showerror("Сохранение", str(e))

    def apply_selected(self) -> None:
        item = self.current_profile()
        if not item: messagebox.showinfo("Профиль", "Выберите профиль."); return
        p = item[1]
        mode = p.get("Mode", "Network")
        mode_label = MODE_LABELS.get(mode, mode)
        device = f"\nУстройство: {p['DeviceName']}" if self.settings.get("show_device", True) and p.get("DeviceName") else ""
        summary = f"Применить «{p.get('Name')}» к адаптеру {self.selected_adapter()}?\n\nТип: {mode_label}\nVLAN: {p.get('VLAN') or 'не используется'}{device}\n\nЭто изменит сетевые параметры Windows."
        if not messagebox.askyesno("Подтверждение применения", summary, icon="warning", parent=self): return
        self.run_action("Применение профиля…", lambda: self.backend.apply(self.selected_adapter(), p), "Профиль применён.", refresh=True,
                        on_success=lambda: self.after(5200, lambda: self.ping_target(p["PingTarget"], automatic=True))
                        if p.get("PingTarget") and self.settings.get("enable_ping", True) else None)

    def run_action(self, pending: str, action, success: str, refresh: bool = False, on_success=None) -> None:
        self.status_var.set(pending)
        def work():
            try:
                action(); logging.info("%s: success", pending)
                def completed():
                    self.status_var.set(success)
                    messagebox.showinfo("Готово", success, parent=self)
                    if on_success: on_success()
                self.after(0, completed)
            except Exception as e:
                error_text = str(e)
                display_error = user_friendly_error(e)
                logging.exception("%s failed", pending)
                self.after(0, lambda: (self.status_var.set("Операция завершилась ошибкой"), messagebox.showerror("Ошибка операции", display_error, parent=self)))
            finally:
                if refresh: self.after(1500, self.refresh_state)
        threading.Thread(target=work, daemon=True).start()

    def restore_dhcp(self) -> None:
        name = self.selected_adapter()
        if messagebox.askyesno("DHCP / DNS", f"Снять VLAN и статические адреса, затем включить DHCP и DNS для {name}?", icon="warning", parent=self):
            self.run_action("Возврат DHCP…", lambda: self.backend.dhcp(name), "DHCP и DNS восстановлены.", True)

    def restore_original(self) -> None:
        name = self.selected_adapter()
        if messagebox.askyesno("Исходные параметры", f"Восстановить снимок сетевого состояния для {name}?", icon="warning", parent=self):
            self.run_action("Восстановление снимка…", lambda: self.backend.restore(name), "Исходные параметры восстановлены.", True)

    def manual_ping(self) -> None:
        target = simpledialog.askstring("Проверка связи", "IP-адрес или имя узла:", parent=self)
        if target: self.ping_target(target)

    def ping_target(self, target: str, automatic: bool = False) -> None:
        def ping():
            try:
                attempts = 5 if automatic else 1
                result = None
                for attempt in range(attempts):
                    result = subprocess.run(["ping.exe", "-n", "2", "-w", "500", target], capture_output=True, text=True, timeout=8)
                    if result.returncode == 0: break
                    if attempt + 1 < attempts: time.sleep(1)
                assert result is not None
                output = result.stdout[-3500:] or result.stderr
                logging.info("Ping %s exit=%s", target, result.returncode)
                message = output + ("\nУзел отвечает." if not result.returncode else "\nОтвет не получен.")
                if not automatic or result.returncode:
                    self.after(0, lambda: messagebox.showinfo(f"Ping · {target}", message, parent=self))
            except Exception as e:
                error_text = str(e)
                self.after(0, lambda: messagebox.showerror("Ping", error_text, parent=self))
        threading.Thread(target=ping, daemon=True).start()

    def app_settings(self) -> None:
        win = tk.Toplevel(self); win.title("Настройки"); win.transient(self); win.grab_set(); win.resizable(False, False)
        ttk.Label(win, text="Параметры приложения").pack(anchor="w", padx=15, pady=(15, 8))
        show = tk.BooleanVar(value=self.settings.get("show_device", True)); ping = tk.BooleanVar(value=self.settings.get("enable_ping", True))
        ttk.Checkbutton(win, text="Показывать DeviceName профиля", variable=show).pack(anchor="w", padx=15, pady=4)
        ttk.Checkbutton(win, text="Автоматически проверять PingTarget", variable=ping).pack(anchor="w", padx=15, pady=4)
        def save():
            self.settings.update(show_device=show.get(), enable_ping=ping.get()); self.save_settings(); win.destroy()
        ttk.Button(win, text="Сохранить", command=save).pack(anchor="e", padx=15, pady=12)


def main() -> None:
    setup_logging()
    if sys.platform != "win32":
        root = tk.Tk(); root.withdraw(); messagebox.showerror(APP_NAME, "Приложение управляет адаптерами Windows и запускается только в Windows."); return
    NetworkManagerApp().mainloop()


if __name__ == "__main__":
    main()
