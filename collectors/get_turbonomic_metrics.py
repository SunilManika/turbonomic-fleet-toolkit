#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Turbonomic VM Metrics Collector
================================
Retrieves an extensive set of metrics for VMs listed in config/config.json from the
Turbonomic REST API and writes JSON, CSV and HTML reports to ./output/.

Metrics collected
-----------------
  Identity & Config     : name, UUID, state, severity, environment type,
                          power state, IP addresses, resource ID
  OS & Software         : OS type, installed IPs, vCPUs
  Cloud tier            : instance type (template), cloud provider / account,
                          region, resource group, billing type, RI coverage %,
                          tenancy, uptime %
  Provisioned config    : vCPUs, memory provisioned (GB), CPU capacity (MHz),
                          memory capacity (GB), num disks, num NICs
  Storage (per-disk)    : storage amount (GiB), IOPS capacity, IO throughput cap
  Current usage         : CPU used (MHz), memory used (GB), IOPS used,
                          net throughput (Kbit/s), IO throughput (Kbit/s),
                          storage latency (ms), storage amount used (MB)
  Peak values           : CPU peak, memory peak, IOPS peak, net throughput peak,
                          IO throughput peak (from values.max in historical stats)
  Avg utilization %     : CPU, memory, IOPS, net throughput, IO throughput,
                          storage latency, storage amount
  Pending Actions       : action type, details, severity, state, risk,
                          current value, new value, estimated savings
  Supply chain          : providers list from entity data (ComputeTier,
                          VirtualVolume, etc.)
  Raw stats             : full flattened commodity map (JSON output only)

Usage
-----
    python get_turbonomic_metrics.py                            # uses ../config/config.json
    python get_turbonomic_metrics.py --config ../config/config.json
    python get_turbonomic_metrics.py --vm turbo-wmi-win-d14
    python get_turbonomic_metrics.py --days 7                  # historical period

Dependencies
------------
    pip install requests urllib3
"""

import argparse
import csv
import html as html_lib
import json
import logging
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Turbonomic API client
# ---------------------------------------------------------------------------

class TurbonomicClient:
    """Turbonomic REST API v3 client - mapped to actual commodity names returned
    by the API (mixed-case, no underscores) confirmed via live introspection."""

    # -----------------------------------------------------------------------
    # Commodity names exactly as Turbonomic returns them (case-sensitive).
    # Confirmed live against https://tz1.demo.turbonomic.com
    # -----------------------------------------------------------------------
    COMMODITY_TYPES = [
        # ── Compute ─────────────────────────────────────────────────────────
        "VCPU",               # VM-level vCPU used (MHz) — sold commodity
        "CPU",                # Host/physical CPU used (MHz) — bought commodity
        "CPUProvisioned",     # CPU provisioned to the VM (MHz)
        "VMem",               # VM memory used (KB) — sold commodity
        "Mem",                # Host memory used (KB) — bought commodity
        "MemProvisioned",     # Memory provisioned to the VM (KB)

        # ── Storage ─────────────────────────────────────────────────────────
        "StorageAmount",      # Storage amount used / provisioned (MB in stats, GiB in disk aspect)
        "StorageAccess",      # IOPS used / capacity
        "StorageLatency",     # Storage latency (ms)
        "StorageProvisioned", # Thin-provisioned / over-provisioned storage (MB)

        # ── Network ─────────────────────────────────────────────────────────
        "NetThroughput",      # Network throughput (Kbit/sec)
        "IOThroughput",       # IO / storage throughput (Kbit/sec)

        # ── Infrastructure counts ────────────────────────────────────────
        "numVCPUs",           # Number of vCPUs (provisioned count)
        "NumDisk",            # Number of disks
        "NetworkInterfaceCount", # Number of NICs

        # ── License / access commodities ────────────────────────────────
        "LicenseAccess",      # License consumption (Windows, RHEL, etc.)
        "TenancyAccess",      # Dedicated / default tenancy commodity

        # ── Core counts (from tier stats) ────────────────────────────────
        "numCores",           # Physical / vCore count for the instance type
        "NumVCore",           # Virtual core count (Azure naming)
    ]

    # -----------------------------------------------------------------------
    # Login endpoints — tried in order, first success wins
    # -----------------------------------------------------------------------
    LOGIN_ENDPOINTS = [
        "/api/v3/login",
        "/vmturbo/rest/login",
    ]

    def __init__(self, url: str, username: str, password: str,
                 verify_ssl: bool = False, timeout: int = 60):
        self.base_url   = url.rstrip("/")
        self.username   = username
        self.password   = password
        self.verify_ssl = verify_ssl
        self.timeout    = timeout
        self.session    = requests.Session()
        self.session.verify = verify_ssl
        self._authenticated = False

    # ------------------------------------------------------------------
    # Authentication
    # ------------------------------------------------------------------

    def authenticate(self) -> None:
        """Try each login endpoint with form-encoded credentials."""
        log.info("Authenticating to %s as %s", self.base_url, self.username)
        payload  = {"username": self.username, "password": self.password}
        last_exc: Exception | None = None

        for endpoint in self.LOGIN_ENDPOINTS:
            url = f"{self.base_url}{endpoint}"
            try:
                resp = self.session.post(
                    url,
                    data=payload,
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    timeout=self.timeout,
                    allow_redirects=True,
                )
                if resp.status_code in (200, 204):
                    self._authenticated = True
                    log.info("Authentication successful via %s", endpoint)
                    return
                if resp.status_code not in (400, 404, 405):
                    resp.raise_for_status()
                log.debug("Endpoint %s returned %s — trying next", endpoint, resp.status_code)
                last_exc = requests.HTTPError(
                    f"{resp.status_code} {resp.reason}", response=resp
                )
            except requests.ConnectionError as exc:
                log.debug("Endpoint %s unreachable: %s", endpoint, exc)
                last_exc = exc

        raise last_exc or RuntimeError("All authentication endpoints failed")

    def _get(self, path: str, params: dict | None = None) -> Any:
        if not self._authenticated:
            self.authenticate()
        url  = f"{self.base_url}{path}"
        resp = self.session.get(url, params=params, timeout=self.timeout)
        if resp.status_code == 401:
            self.authenticate()
            resp = self.session.get(url, params=params, timeout=self.timeout)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        text = resp.text.strip()
        return resp.json() if text else None

    def _post(self, path: str, body: dict, params: dict | None = None) -> Any:
        if not self._authenticated:
            self.authenticate()
        url  = f"{self.base_url}{path}"
        resp = self.session.post(url, json=body, params=params, timeout=self.timeout)
        if resp.status_code == 401:
            self.authenticate()
            resp = self.session.post(url, json=body, params=params, timeout=self.timeout)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        text = resp.text.strip()
        return resp.json() if text else None

    # ------------------------------------------------------------------
    # Entity search
    # ------------------------------------------------------------------

    def search_vms(self, vm_names: list[str]) -> list[dict]:
        """Search for VirtualMachine entities by display name."""
        log.info("Searching for %d VMs in Turbonomic", len(vm_names))
        body = {
            "criteriaList": [
                {
                    "filterType":    "vmsByName",
                    "expType":       "RXEQ",
                    "expVal":        "|".join(vm_names),
                    "caseSensitive": False,
                }
            ],
            "logicalOperator": "AND",
            "className":       "VirtualMachine",
        }
        try:
            results = self._post("/api/v3/search", body=body,
                                 params={"types": "VirtualMachine"})
            if not results:
                return []
            log.info("Found %d VM entities", len(results))
            return results
        except Exception as exc:
            log.warning("Search failed: %s — falling back to full entity list", exc)
            return self._get_all_vms(vm_names)

    def _get_all_vms(self, vm_names: list[str]) -> list[dict]:
        log.info("Fetching full VM list from Turbonomic")
        all_vms = self._get("/api/v3/entities",
                            params={"types": "VirtualMachine", "limit": 5000})
        if not all_vms:
            return []
        names_lower = {n.lower() for n in vm_names}
        matched = [e for e in all_vms
                   if e.get("displayName", "").lower() in names_lower]
        log.info("Matched %d VMs by name from full list", len(matched))
        return matched

    # ------------------------------------------------------------------
    # Entity data fetchers
    # ------------------------------------------------------------------

    def get_entity_details(self, uuid: str) -> dict:
        try:
            result = self._get(f"/api/v3/entities/{uuid}")
            return result or {}
        except Exception as exc:
            log.warning("Details unavailable for %s: %s", uuid, exc)
            return {}

    def get_tier_stats(self, tier_uuid: str) -> dict[str, Any]:
        """Fetch capacity specs for a ComputeTier entity (instance type sizing)."""
        try:
            result = self._post(f"/api/v3/entities/{tier_uuid}/stats", body={})
            if not isinstance(result, list) or not result:
                return {}
            # Flatten: name → capacity.avg
            caps: dict[str, Any] = {}
            for stat in result[0].get("statistics", []):
                name = stat.get("name") or ""
                cap_block = stat.get("capacity") or {}
                cap_v = _safe_num(cap_block.get("avg") if isinstance(cap_block, dict)
                                  else stat.get("capacity"))
                if name and cap_v is not None:
                    caps[name] = cap_v
            return caps
        except Exception as exc:
            log.debug("Tier stats unavailable for %s: %s", tier_uuid, exc)
            return {}

    def get_entity_aspects(self, uuid: str) -> dict:
        try:
            result = self._get(f"/api/v3/entities/{uuid}/aspects")
            return result or {}
        except Exception as exc:
            log.debug("Aspects unavailable for %s: %s", uuid, exc)
            return {}

    def get_entity_stats(self, uuid: str, days: int = 0) -> list[dict]:
        """
        Fetch statistics for all commodity types.
        days=0  → latest snapshot only (value = current reading).
        days>0  → historical window; values.avg = average, values.max = peak.
        Always request with days>=1 to get peak data in values.max.
        Uses minimum 1 day so peak is always populated.
        """
        # Always use at least 1 day so we get the values.max (peak) field
        effective_days = max(days, 1)
        now   = datetime.now(timezone.utc)
        start = now - timedelta(days=effective_days)
        body: dict = {
            "statistics": [{"name": c} for c in self.COMMODITY_TYPES],
            "startDate":  int(start.timestamp() * 1000),
            "endDate":    int(now.timestamp() * 1000),
        }
        try:
            result = self._post(f"/api/v3/entities/{uuid}/stats", body=body)
            return result if isinstance(result, list) else []
        except Exception as exc:
            log.warning("Stats unavailable for %s: %s", uuid, exc)
            return []

    def get_entity_actions(self, uuid: str) -> list[dict]:
        try:
            result = self._get(f"/api/v3/entities/{uuid}/actions",
                               params={"limit": 500})
            return result if isinstance(result, list) else []
        except Exception as exc:
            log.debug("Actions unavailable for %s: %s", uuid, exc)
            return []

    def get_entity_tags(self, uuid: str) -> dict:
        try:
            result = self._get(f"/api/v3/entities/{uuid}/tags")
            return result if isinstance(result, dict) else {}
        except Exception as exc:
            log.debug("Tags unavailable for %s: %s", uuid, exc)
            return {}

    # ------------------------------------------------------------------
    # Stats flattening
    # ------------------------------------------------------------------

    @staticmethod
    def _flatten_stats(snapshots: list[dict]) -> dict[str, Any]:
        """
        Merge all stat snapshots into one dict keyed by commodity name.
        Uses the last snapshot's values (most recent), but tracks overall max
        across all snapshots for peak.

        The Turbonomic API for this instance returns:
          - stat.value        → current / avg value
          - stat.values.avg   → period average
          - stat.values.max   → period peak
          - stat.values.min   → period minimum
          - stat.capacity.avg → capacity
          - stat.relation     → None (no bought/sold distinction here)
          - stat.units        → unit string
        """
        # Build: name → accumulated info across all snapshots
        acc: dict[str, dict] = {}

        for snapshot in snapshots:
            for stat in snapshot.get("statistics", []):
                name = stat.get("name") or "UNKNOWN"

                val_block = stat.get("values") or {}
                cap_block = stat.get("capacity") or {}

                avg_v = _safe_num(val_block.get("avg") if isinstance(val_block, dict)
                                  else stat.get("value"))
                max_v = _safe_num(val_block.get("max") if isinstance(val_block, dict)
                                  else None)
                min_v = _safe_num(val_block.get("min") if isinstance(val_block, dict)
                                  else None)
                cap_v = _safe_num(cap_block.get("avg") if isinstance(cap_block, dict)
                                  else stat.get("capacity"))

                if name not in acc:
                    acc[name] = {
                        "commodity":       name,
                        "relation":        stat.get("relation") or "",
                        "units":           stat.get("units") or "",
                        "used":            avg_v,
                        "used_avg":        avg_v,
                        "used_min":        min_v,
                        "used_max":        max_v,      # peak
                        "capacity":        cap_v,
                        "peak":            max_v,
                        "utilization_pct": _util_pct(avg_v, cap_v),
                        "displayName":     stat.get("displayName") or "",
                    }
                else:
                    # Keep max peak across all snapshots
                    existing_peak = acc[name]["peak"]
                    if max_v is not None and (existing_peak is None or max_v > existing_peak):
                        acc[name]["peak"]     = max_v
                        acc[name]["used_max"] = max_v
                    # Update avg to latest snapshot value
                    acc[name]["used"]     = avg_v
                    acc[name]["used_avg"] = avg_v

        return acc

    # ------------------------------------------------------------------
    # Full metric collection per VM
    # ------------------------------------------------------------------

    def collect_vm_metrics(self, entity: dict, days: int = 0) -> dict:
        """Collect every available metric for one VM entity."""
        uuid = entity.get("uuid", "")
        name = entity.get("displayName", uuid)
        log.info("  Collecting: %s (%s)", name, uuid)

        details  = self.get_entity_details(uuid)
        aspects  = self.get_entity_aspects(uuid)
        stats    = self.get_entity_stats(uuid, days=days)
        actions  = self.get_entity_actions(uuid)
        tags     = self.get_entity_tags(uuid)

        vm = details if details else entity

        # ── Aspects ──────────────────────────────────────────────────────
        av       = aspects.get("virtualMachineAspect", {}) if aspects else {}
        cloud_av = aspects.get("cloudAspect", {}) if aspects else {}
        disks_av = aspects.get("virtualDisksAspect", {}) if aspects else {}

        stats_flat = self._flatten_stats(stats)

        # ── Convenience helpers ───────────────────────────────────────────
        def sv(commodity: str, field: str = "used") -> Any:
            """Return a named field from a flattened commodity entry."""
            entry = stats_flat.get(commodity)
            return entry.get(field) if entry else None

        def su(commodity: str) -> Any:
            """Return utilization_pct for a commodity."""
            return sv(commodity, "utilization_pct")

        def spk(commodity: str) -> Any:
            """Return peak (values.max) for a commodity."""
            return sv(commodity, "peak")

        def sc(commodity: str) -> Any:
            """Return capacity for a commodity."""
            return sv(commodity, "capacity")

        # ── OS & Identity ─────────────────────────────────────────────────
        # virtualMachineAspect has: os, ip (list), numVCPUs, resourceId
        os_type   = _first_val(av.get("os"), vm.get("guestOsType"), "")
        ip_list   = av.get("ip") or vm.get("ipAddress") or []
        num_vcpus = _first_val(av.get("numVCPUs"), vm.get("numCPUs"),
                               int(sv("numVCPUs") or 0) or None)
        resource_id = av.get("resourceId") or cloud_av.get("resourceId") or ""

        # ── Cloud tier ────────────────────────────────────────────────────
        # cloudAspect.template.displayName = instance type (e.g. Standard_B4as_v2)
        template_obj = cloud_av.get("template") or {}
        cloud_tier   = _first_val(
            template_obj.get("displayName") if isinstance(template_obj, dict) else None,
            vm.get("templateName"), ""
        )

        # cloudAspect.region is a dict {uuid, displayName, className}
        region_obj = cloud_av.get("region") or {}
        region     = (region_obj.get("displayName") if isinstance(region_obj, dict)
                      else str(region_obj)) or ""

        # cloudAspect.businessAccount.displayName
        account_obj  = cloud_av.get("businessAccount") or {}
        account_name = (account_obj.get("displayName") if isinstance(account_obj, dict)
                        else str(account_obj)) or ""
        account_id   = (account_obj.get("uuid") if isinstance(account_obj, dict)
                        else "") or ""

        # cloudAspect.resourceGroup.displayName
        rg_obj = cloud_av.get("resourceGroup") or {}
        resource_group = (rg_obj.get("displayName") if isinstance(rg_obj, dict)
                          else str(rg_obj)) or ""

        cloud_provider = _first_val(
            cloud_av.get("cloudProvider"), vm.get("environmentType"), ""
        )
        billing_type      = cloud_av.get("billingType", "")
        ri_coverage_pct   = _safe_num(cloud_av.get("riCoveragePercentage"))
        tenancy           = cloud_av.get("tenancy", "")
        cloud_service     = cloud_av.get("cloudServiceName", "")

        uptime_obj      = cloud_av.get("entityUptime") or {}
        uptime_pct      = _safe_num(uptime_obj.get("uptimePercentage")) \
                          if isinstance(uptime_obj, dict) else None

        # ── Provisioned — from stats (CPUProvisioned / MemProvisioned) ────
        cpu_prov_mhz  = sv("CPUProvisioned")
        mem_prov_kb   = sv("MemProvisioned")
        cpu_cap_mhz   = _first_val(sc("VCPU"), sc("CPU"), cpu_prov_mhz)
        mem_cap_kb    = _first_val(sc("VMem"), sc("Mem"), mem_prov_kb)
        num_disks     = _safe_num(sc("NumDisk")) or _safe_num(sv("NumDisk"))
        num_nics      = _safe_num(sc("NetworkInterfaceCount")) or \
                        _safe_num(sv("NetworkInterfaceCount"))

        # ── Disk stats from virtualDisksAspect ───────────────────────────
        disk_list = disks_av.get("virtualDisks", [])
        disks_info = []
        total_storage_gib  = 0.0
        total_iops_cap     = 0.0
        total_io_cap_kbps  = 0.0
        total_iops_used    = 0.0
        total_io_used_kbps = 0.0
        for disk in disk_list:
            d_stats = {s.get("name"): s for s in disk.get("stats", [])}
            sa   = d_stats.get("StorageAmount", {})
            acc  = d_stats.get("StorageAccess", {})
            iot  = d_stats.get("IOThroughput", {})
            sa_cap  = _safe_num((sa.get("capacity") or {}).get("avg"))
            acc_cap = _safe_num((acc.get("capacity") or {}).get("avg"))
            iot_cap = _safe_num((iot.get("capacity") or {}).get("avg"))
            acc_use = _safe_num(acc.get("value"))
            iot_use = _safe_num(iot.get("value"))
            if sa_cap:
                total_storage_gib  += sa_cap
            if acc_cap:
                total_iops_cap     += acc_cap
            if iot_cap:
                total_io_cap_kbps  += iot_cap
            if acc_use:
                total_iops_used    += acc_use
            if iot_use:
                total_io_used_kbps += iot_use
            disks_info.append({
                "name":           disk.get("displayName", ""),
                "tier":           disk.get("tier", ""),
                "storageAmtGiB":  sa_cap,
                "iopsCap":        acc_cap,
                "iopsUsed":       acc_use,
                "ioThroughputCapKBps": iot_cap,
                "ioThroughputUsedKBps": iot_use,
            })

        # Roll-up disk totals (use disk aspect values as ground truth for storage)
        storage_prov_gib  = total_storage_gib if total_storage_gib > 0 else None
        iops_cap_total    = total_iops_cap    if total_iops_cap > 0 else None
        iops_used_total   = total_iops_used   if total_iops_used > 0 else \
                            sv("StorageAccess")
        io_cap_kbps       = total_io_cap_kbps if total_io_cap_kbps > 0 else None
        io_used_kbps      = total_io_used_kbps if total_io_used_kbps > 0 else None

        # ── Providers from entity data (not a separate API call) ──────────
        providers_raw = vm.get("providers") or entity.get("providers") or []
        providers_flat = [
            {
                "uuid":  p.get("uuid", ""),
                "name":  p.get("displayName", ""),
                "type":  p.get("className", ""),
            }
            for p in providers_raw
        ]
        compute_tier_name = next(
            (p["name"] for p in providers_flat if p["type"] == "ComputeTier"), "")
        storage_name = next(
            (p["name"] for p in providers_flat
             if p["type"] in ("VirtualVolume", "Storage")), "")

        # ── Tags ──────────────────────────────────────────────────────────
        tags_flat: dict = {}
        for k, vals in (tags or {}).items():
            tags_flat[k] = ", ".join(vals) if isinstance(vals, list) else str(vals)

        # ── Consumers ────────────────────────────────────────────────────
        consumers_raw = vm.get("consumers") or entity.get("consumers") or []
        consumers_flat = [
            {"uuid": c.get("uuid",""), "name": c.get("displayName",""),
             "type": c.get("className","")}
            for c in consumers_raw
        ]

        # ── Actions ───────────────────────────────────────────────────────
        actions_flat = []
        total_savings_per_hour = 0.0
        for a in actions:
            tgt          = a.get("target") or {}
            risk_obj     = a.get("risk") or {}
            exec_chars   = a.get("executionCharacteristics") or {}
            cur_ent      = a.get("currentEntity") or {}
            new_ent      = a.get("newEntity") or {}
            cur_loc      = a.get("currentLocation") or {}
            new_loc      = a.get("newLocation") or {}

            # savings $/h lives in action.stats[] with filter savingsType=savings
            savings_per_hour = None
            for s in (a.get("stats") or []):
                filters = s.get("filters") or []
                if any(f.get("value") == "savings" for f in filters if isinstance(f, dict)):
                    savings_per_hour = _safe_num(s.get("value"))
                    break
            if savings_per_hour:
                total_savings_per_hour += savings_per_hour

            actions_flat.append({
                "actionType":         a.get("actionType", ""),
                "actionMode":         a.get("actionMode", ""),
                "actionStateDesc":    a.get("actionStateDescription", ""),
                "details":            a.get("details", ""),
                "severity":           a.get("severity", ""),
                "state":              a.get("actionState", ""),
                "risk":               risk_obj.get("description", ""),
                "riskSubCategory":    risk_obj.get("subCategory", ""),
                "riskSeverity":       risk_obj.get("severity", ""),
                "reasonCommodities":  ", ".join(risk_obj.get("reasonCommodities") or []),
                "disruptiveness":     exec_chars.get("disruptiveness", ""),
                "reversibility":      exec_chars.get("reversibility", ""),
                "savingsPerHour":     savings_per_hour,
                "savingsPerMonth":    round(savings_per_hour * 730, 4) if savings_per_hour else None,
                "target":             tgt.get("displayName", ""),
                "target_uuid":        tgt.get("uuid", ""),
                "currentEntity":      cur_ent.get("displayName", ""),
                "currentEntityUuid":  cur_ent.get("uuid", ""),
                "newEntity":          new_ent.get("displayName", ""),
                "newEntityUuid":      new_ent.get("uuid", ""),
                "currentLocation":    cur_loc.get("displayName", ""),
                "newLocation":        new_loc.get("displayName", ""),
                "importance":         _safe_num(a.get("importance")),
                "createTime":         a.get("createTime", ""),
                # Legacy fields kept for backward compatibility
                "newValue":           str(a.get("newValue", "")),
                "currentValue":       str(a.get("currentValue", "")),
            })

        # ── Compute tier specs (current + recommended from actions) ───────
        # Fetch specs for current tier — zero extra auth, one GET per VM
        tier_obj     = vm.get("template") or {}
        tier_uuid    = tier_obj.get("uuid", "")
        tier_price   = _safe_num(tier_obj.get("price"))        # on-demand $/h
        tier_specs   = self.get_tier_stats(tier_uuid) if tier_uuid else {}
        tier_family  = ""
        tier_aspects = {}
        if tier_uuid:
            try:
                ta = self._get(f"/api/v3/entities/{tier_uuid}/aspects")
                if ta:
                    tier_aspects = ta.get("computeTierAspect") or {}
                    tier_family  = tier_aspects.get("tierFamily", "")
            except Exception:
                pass

        # Physical / virtual cores from tier stats
        # numCores = physical cores per socket for the instance type (cloud = same as vCores)
        # NumVCore = virtual core count (Azure terminology)
        tier_num_cores  = tier_specs.get("numCores") or tier_specs.get("NumVCore")
        # Hyper-thread ratio = vCPUs / physical cores
        # For cloud VMs there is no physical host exposed; the ratio reflects
        # how many vCPUs are mapped onto each logical core of the instance type.
        ht_ratio: float | None = None
        if tier_num_cores and num_vcpus and float(tier_num_cores) > 0:
            ht_ratio = round(float(num_vcpus) / float(tier_num_cores), 2)

        # Recommended tier comes from the first SCALE action's newEntity
        rec_tier_name = ""
        rec_tier_uuid = ""
        rec_tier_specs: dict = {}
        for a_flat in actions_flat:
            if a_flat.get("actionType") == "SCALE" and a_flat.get("newEntity"):
                rec_tier_name = a_flat["newEntity"]
                rec_tier_uuid = a_flat["newEntityUuid"]
                if rec_tier_uuid:
                    rec_tier_specs = self.get_tier_stats(rec_tier_uuid)
                break

        # ── VM cost from entity ──────────────────────────────────────────
        vm_cost_per_hour = _safe_num(vm.get("costPrice"))

        # ── Compute utilization % (direct calc for VCPU / VMem) ──────────
        vcpu_util_pct = _util_pct(sv("VCPU"), sc("VCPU") or cpu_cap_mhz)
        vmem_util_pct = _util_pct(sv("VMem"), sc("VMem") or mem_cap_kb)

        return {
            # ── Identity ─────────────────────────────────────────────────
            "uuid":             uuid,
            "name":             name,
            "state":            vm.get("state", ""),
            "severity":         vm.get("severity", ""),
            "severityBreakdown": vm.get("severityBreakdown") or {},
            "staleness":        vm.get("staleness", ""),
            "powerState":       vm.get("powerState", ""),
            "environmentType":  vm.get("environmentType", ""),
            "ipAddresses":      ", ".join(str(ip) for ip in ip_list) if ip_list else "",
            "resourceId":       resource_id,
            "vendorId":         list((vm.get("vendorIds") or {}).values())[0]
                                if vm.get("vendorIds") else "",
            "discoveredByType": (vm.get("discoveredBy") or {}).get("type", ""),
            "discoveredByName": (vm.get("discoveredBy") or {}).get("displayName", ""),

            # ── OS & Software ─────────────────────────────────────────────
            "osType":           os_type,
            "osVersion":        "",           # Not provided by this Turbonomic instance

            # ── Cloud / Infrastructure tier ──────────────────────────────
            "cloudTier":        cloud_tier,       # e.g. Standard_B4as_v2
            "cloudProvider":    cloud_provider,   # e.g. CLOUD / Azure
            "cloudService":     cloud_service,    # e.g. AZURE_VIRTUAL_MACHINES
            "region":           region,           # e.g. azure-East US
            "resourceGroup":    resource_group,   # e.g. rg-xxxx
            "accountName":      account_name,     # e.g. itz-custom-requests
            "accountId":        account_id,
            "billingType":      billing_type,     # e.g. ONDEMAND
            "riCoveragePct":    ri_coverage_pct,
            "tenancy":          tenancy,          # e.g. DEFAULT
            "uptimePct":        uptime_pct,
            "computeTierName":  compute_tier_name,
            "storageName":      storage_name,
            "tags":             tags_flat,

            # ── Cost ─────────────────────────────────────────────────────
            "vmCostPerHour":        vm_cost_per_hour,
            "vmCostPerMonth":       round(vm_cost_per_hour * 730, 4) if vm_cost_per_hour else None,
            "tierOnDemandPricePerHour": tier_price,
            "tierOnDemandPricePerMonth": round(tier_price * 730, 4) if tier_price else None,
            "totalActionSavingsPerHour":  round(total_savings_per_hour, 4) if total_savings_per_hour else None,
            "totalActionSavingsPerMonth": round(total_savings_per_hour * 730, 4) if total_savings_per_hour else None,

            # ── Provisioned configuration ────────────────────────────────
            "numCPUs":              num_vcpus,
            "cpuCapacityMHz":       cpu_cap_mhz,
            "cpuProvisionedMHz":    cpu_prov_mhz,
            "memCapacityGB":        _kb_to_gb(mem_cap_kb),
            "memProvisionedGB":     _kb_to_gb(mem_prov_kb),
            "numDisks":             len(disk_list) or (int(num_disks) if num_disks else None),
            "numNICs":              int(num_nics) if num_nics else None,
            "storageProvisionedGiB": storage_prov_gib,   # from disk aspect (GiB)
            "iopsCapacity":         iops_cap_total,
            "ioThroughputCapKBps":  io_cap_kbps,

            # ── Physical / core topology ─────────────────────────────────
            # physicalCores: logical cores per instance type (numCores from tier stats).
            # For cloud VMs there is no physical server exposed — this is the vCore
            # count of the instance type, which maps 1:1 to physical cores available.
            "physicalCores":        tier_num_cores,
            # physicalProcessors: Turbonomic does not expose physical socket/processor
            # count for cloud VMs (Azure/AWS abstract the host entirely). Value is None.
            "physicalProcessors":   None,
            # hyperThreadRatio: vCPUs assigned to the VM divided by physicalCores.
            # >1 means hyperthreading is in use (e.g. 8 vCPUs / 4 cores = 2.0).
            "hyperThreadRatio":     ht_ratio,

            # ── Current compute tier specs ────────────────────────────────
            "tierName":             tier_obj.get("displayName", ""),
            "tierFamily":           tier_family,
            "tierCpuCapMHz":        tier_specs.get("CPU") or tier_specs.get("CPUProvisioned"),
            "tierMemCapGB":         _kb_to_gb(tier_specs.get("Mem") or tier_specs.get("MemProvisioned")),
            "tierIopsCapacity":     tier_specs.get("StorageAccess"),
            "tierNetCapKbitps":     tier_specs.get("NetThroughput"),
            "tierIoCapKbitps":      tier_specs.get("IOThroughput"),
            "tierNumVCores":        tier_specs.get("NumVCore") or tier_specs.get("numCores"),
            "tierMaxDisks":         tier_specs.get("NumDisk"),
            "tierMaxNICs":          tier_specs.get("NetworkInterfaceCount"),

            # ── Recommended tier specs (from first SCALE action) ──────────
            "recTierName":          rec_tier_name,
            "recTierCpuCapMHz":     rec_tier_specs.get("CPU") or rec_tier_specs.get("CPUProvisioned"),
            "recTierMemCapGB":      _kb_to_gb(rec_tier_specs.get("Mem") or rec_tier_specs.get("MemProvisioned")),
            "recTierIopsCapacity":  rec_tier_specs.get("StorageAccess"),
            "recTierNetCapKbitps":  rec_tier_specs.get("NetThroughput"),

            # ── Current usage ─────────────────────────────────────────────
            "cpuUsedMHz":           sv("VCPU"),
            "cpuHostUsedMHz":       sv("CPU"),
            "memUsedGB":            _kb_to_gb(sv("VMem")),
            "memHostUsedGB":        _kb_to_gb(sv("Mem")),
            "storageAmountUsedMB":  sv("StorageAmount"),
            "storageAmountCapMB":   sc("StorageAmount"),
            "storageAmountUsedGB":  _mb_to_gb(sv("StorageAmount")),
            "storageAmountCapGB":   _mb_to_gb(sc("StorageAmount")),
            "storageProvUsedMB":    sv("StorageProvisioned"),
            "storageProvCapMB":     sc("StorageProvisioned"),
            "iopsUsed":             iops_used_total,
            "ioThroughputUsedKBps": io_used_kbps,
            # Throughput in MB/s (1 KB/s = 1/1024 MB/s; IOThroughput is in Kbit/s so /8/1024 = MB/s)
            "ioThroughputUsedMBps": round(io_used_kbps / 1024, 3) if io_used_kbps else None,
            "ioThroughputCapMBps":  round(io_cap_kbps / 1024, 3) if io_cap_kbps else None,
            "netThroughputKbitps":  sv("NetThroughput"),
            "netThroughputCapKbitps": sc("NetThroughput"),
            "ioThroughputStatKbitps": sv("IOThroughput"),
            "ioThroughputStatCapKbitps": sc("IOThroughput"),
            # IOThroughput stat is in Kbit/s → convert to MB/s (÷ 8 ÷ 1024)
            "ioThroughputStatMBps": round(float(sv("IOThroughput")) / 8 / 1024, 3)
                                    if sv("IOThroughput") else None,
            "storageLatencyMs":     sv("StorageLatency"),
            "storageLatencyCapMs":  sc("StorageLatency"),
            "licenseAccess":        sv("LicenseAccess"),
            "tenancyAccess":        sv("TenancyAccess"),

            # ── Average utilization % ─────────────────────────────────────
            "cpuUtilizationPct":        vcpu_util_pct,
            "cpuHostUtilPct":           _util_pct(sv("CPU"), sc("CPU")),
            "memUtilizationPct":        vmem_util_pct,
            "memHostUtilPct":           _util_pct(sv("Mem"), sc("Mem")),
            "storageAmountUtilPct":     su("StorageAmount"),
            "storageAccessUtilPct":     su("StorageAccess"),
            "storageLatencyUtilPct":    su("StorageLatency"),
            "netThroughputUtilPct":     su("NetThroughput"),
            "ioThroughputUtilPct":      su("IOThroughput"),
            "iopsUtilPct":              _util_pct(iops_used_total, iops_cap_total),
            "ioThroughputDiskUtilPct":  _util_pct(io_used_kbps, io_cap_kbps),

            # ── Peak values (values.max from historical stats) ────────────
            "cpuPeakMHz":           spk("VCPU") or spk("CPU"),
            "cpuHostPeakMHz":       spk("CPU"),
            # CPU peak utilization % = peak MHz / capacity MHz × 100
            "cpuPeakUtilPct":       _util_pct(spk("VCPU") or spk("CPU"), cpu_cap_mhz),
            "memPeakGB":            _kb_to_gb(spk("VMem") or spk("Mem")),
            "memHostPeakGB":        _kb_to_gb(spk("Mem")),
            # Memory peak utilization % = peak KB / capacity KB × 100
            "memPeakUtilPct":       _util_pct(spk("VMem") or spk("Mem"), mem_cap_kb),
            "storageAmountPeakMB":  spk("StorageAmount"),
            "iopsPeak":             spk("StorageAccess"),
            "netThroughputPeakKbitps": spk("NetThroughput"),
            "ioThroughputPeakKbitps":  spk("IOThroughput"),
            "storageLatencyPeakMs": spk("StorageLatency"),

            # ── Actions ───────────────────────────────────────────────────
            "actionCount":          len(actions_flat),
            "actions":              actions_flat,
            "totalActionSavingsPerHour":  round(total_savings_per_hour, 4) if total_savings_per_hour else None,
            "totalActionSavingsPerMonth": round(total_savings_per_hour * 730, 4) if total_savings_per_hour else None,

            # ── Providers (supply chain from entity data) ─────────────────
            "providers":            providers_flat,

            # ── Consumers ────────────────────────────────────────────────
            "consumerCount":        len(consumers_flat),
            "consumers":            consumers_flat,

            # ── Per-disk detail ───────────────────────────────────────────
            "disks":                disks_info,

            # ── Raw commodity stats ───────────────────────────────────────
            # Full flattened map — only in JSON output
            "stats":                stats_flat,
        }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_num(v: Any) -> float | None:
    try:
        return round(float(v), 4) if v is not None else None
    except (TypeError, ValueError):
        return None


def _util_pct(used: Any, capacity: Any) -> float | None:
    try:
        u = float(used)
        c = float(capacity)
        if c > 0:
            return round(min(100.0, u / c * 100), 2)
    except (TypeError, ValueError):
        pass
    return None


def _mb_to_gb(mb: Any) -> float | None:
    """Convert MB → GB."""
    try:
        return round(float(mb) / 1024, 2) if mb else None
    except (TypeError, ValueError):
        return None


def _kb_to_gb(kb: Any) -> float | None:
    """Convert KB → GB."""
    try:
        return round(float(kb) / (1024 * 1024), 2) if kb else None
    except (TypeError, ValueError):
        return None


def _first_val(*args: Any) -> Any:
    """Return first non-None, non-empty value from args."""
    for a in args:
        if a is not None and a != "" and a != []:
            return a
    return args[-1] if args else None


def extract_vm_names_from_config(cfg: dict, config_dir: Path) -> list[str]:
    """Extract VM names from config dictionary or fallback CSV.

    Sources checked in order:
      1. 'targets.windows' and 'targets.linux' (vm_name, display_name, or nameOrAddress)
      2. 'windows' / 'linux' / 'wmi_targets' lists
      3. 'vms' or 'servers' lists in config
      4. 'servers_csv' referenced CSV file
    """
    names: list[str] = []

    # 1. targets.windows / targets.linux (matching config/config.json structure)
    targets = cfg.get("targets", {})
    if isinstance(targets, dict):
        for target_group in ["windows", "linux"]:
            group_list = targets.get(target_group, [])
            if isinstance(group_list, list):
                for item in group_list:
                    if isinstance(item, dict):
                        n = item.get("vm_name") or item.get("display_name") or item.get("nameOrAddress")
                        if n and str(n).strip() and str(n).strip() not in names:
                            names.append(str(n).strip())
            elif isinstance(group_list, dict):
                n = group_list.get("vm_name") or group_list.get("display_name")
                if n and str(n).strip() and str(n).strip() not in names:
                    names.append(str(n).strip())

    # 2. top-level windows / linux lists
    for top_key in ["windows", "linux", "wmi_targets", "windows_targets"]:
        top_list = cfg.get(top_key, [])
        if isinstance(top_list, list):
            for item in top_list:
                if isinstance(item, dict):
                    n = item.get("vm_name") or item.get("display_name") or item.get("nameOrAddress")
                    if n and str(n).strip() and str(n).strip() not in names:
                        names.append(str(n).strip())

    # 3. direct list of VM names
    if not names and "vms" in cfg and isinstance(cfg["vms"], list):
        for item in cfg["vms"]:
            if isinstance(item, str) and item.strip() and item.strip() not in names:
                names.append(item.strip())
            elif isinstance(item, dict):
                n = item.get("vm_name") or item.get("name")
                if n and str(n).strip() and str(n).strip() not in names:
                    names.append(str(n).strip())

    # 4. fallback to servers_csv if provided and exists
    if not names and "servers_csv" in cfg:
        csv_rel = cfg["servers_csv"]
        csv_path = (config_dir / csv_rel).resolve()
        if csv_path.exists():
            names = load_vm_names_from_csv(str(csv_path))

    return names


def load_vm_names_from_csv(csv_path: str) -> list[str]:
    names = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            name = (row.get("Name") or row.get("name") or row.get("vm_name") or "").strip()
            if name:
                names.append(name)
    log.info("Loaded %d VM names from %s", len(names), csv_path)
    return names


# ---------------------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------------------

def write_json(data: list[dict], path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=str)
    log.info("JSON written: %s", path)


# All scalar columns for CSV (list/dict fields excluded)
CSV_COLUMNS = [
    # Identity
    "name", "uuid", "state", "severity", "staleness", "powerState", "environmentType",
    "ipAddresses", "resourceId", "vendorId", "discoveredByType", "discoveredByName",
    # OS
    "osType",
    # Cloud tier
    "cloudTier", "cloudProvider", "cloudService", "region", "resourceGroup",
    "accountName", "accountId", "billingType", "riCoveragePct", "tenancy",
    "uptimePct", "computeTierName", "storageName",
    # Cost
    "vmCostPerHour", "vmCostPerMonth",
    "tierOnDemandPricePerHour", "tierOnDemandPricePerMonth",
    "totalActionSavingsPerHour", "totalActionSavingsPerMonth",
    # Provisioned config
    "numCPUs", "cpuCapacityMHz", "cpuProvisionedMHz",
    "memCapacityGB", "memProvisionedGB",
    "numDisks", "numNICs",
    "storageProvisionedGiB", "iopsCapacity", "ioThroughputCapKBps",
    # Physical / core topology
    "physicalCores", "physicalProcessors", "hyperThreadRatio",
    # Current tier specs
    "tierName", "tierFamily",
    "tierCpuCapMHz", "tierMemCapGB", "tierIopsCapacity",
    "tierNetCapKbitps", "tierIoCapKbitps", "tierNumVCores", "tierMaxDisks", "tierMaxNICs",
    # Recommended tier specs
    "recTierName",
    "recTierCpuCapMHz", "recTierMemCapGB", "recTierIopsCapacity", "recTierNetCapKbitps",
    # Current usage
    "cpuUsedMHz", "cpuHostUsedMHz",
    "memUsedGB", "memHostUsedGB",
    "storageAmountUsedGB", "storageAmountCapGB",
    "storageProvUsedMB", "storageProvCapMB",
    "iopsUsed", "ioThroughputUsedKBps", "ioThroughputUsedMBps", "ioThroughputCapMBps",
    "netThroughputKbitps", "netThroughputCapKbitps",
    "ioThroughputStatKbitps", "ioThroughputStatCapKbitps", "ioThroughputStatMBps",
    "storageLatencyMs", "storageLatencyCapMs",
    # Avg utilization %
    "cpuUtilizationPct", "cpuHostUtilPct",
    "memUtilizationPct", "memHostUtilPct",
    "storageAmountUtilPct", "storageAccessUtilPct", "storageLatencyUtilPct",
    "netThroughputUtilPct", "ioThroughputUtilPct",
    "iopsUtilPct", "ioThroughputDiskUtilPct",
    # Peak + peak utilization %
    "cpuPeakMHz", "cpuHostPeakMHz", "cpuPeakUtilPct",
    "memPeakGB", "memHostPeakGB", "memPeakUtilPct",
    "storageAmountPeakMB", "iopsPeak",
    "netThroughputPeakKbitps", "ioThroughputPeakKbitps", "storageLatencyPeakMs",
    # Consumers / actions
    "consumerCount", "actionCount",
]


def write_csv(data: list[dict], path: str) -> None:
    if not data:
        return
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for vm in data:
            writer.writerow({k: vm.get(k, "") for k in CSV_COLUMNS})
    log.info("CSV written: %s", path)


def write_html(data: list[dict], path: str, generated_at: str) -> None:
    """HTML dashboard — fully redesigned with sidebar nav, badges, progress bars and SVG charts."""

    def enc(v: Any) -> str:
        return html_lib.escape(str(v)) if v is not None and v != "" else "—"

    def pct_bar(pct: Any) -> str:
        try:
            p = float(pct)
        except (TypeError, ValueError):
            return "<span class='muted'>—</span>"
        color = "var(--ok)" if p < 70 else "var(--warn)" if p < 90 else "var(--bad)"
        return (
            f"<div class='pbar-wrap'>"
            f"<div class='pbar-track'><div class='pbar-fill' style='width:{min(p,100):.1f}%;background:{color}'></div></div>"
            f"<span class='pbar-label'>{p:.1f}%</span></div>"
        )

    def badge(text: str, kind: str = "muted") -> str:
        return f"<span class='badge badge-{kind}'>{html_lib.escape(str(text))}</span>"

    def sev_badge(s: str) -> str:
        k = {"CRITICAL": "bad", "MAJOR": "warn", "MINOR": "warn", "NORMAL": "ok"}.get(
            str(s).upper(), "muted")
        return badge(s, k)

    def power_badge(s: str) -> str:
        return badge(s, "ok" if str(s).upper() == "POWERED_ON" else "muted")

    def gb(v: Any) -> str:
        try:
            return f"{float(v):.2f}"
        except (TypeError, ValueError):
            return "—"

    def usd(v: Any) -> str:
        try:
            return f"${float(v):.4f}"
        except (TypeError, ValueError):
            return "—"

    def usd2(v: Any) -> str:
        try:
            return f"${float(v):.2f}"
        except (TypeError, ValueError):
            return "—"

    def n(v: Any) -> str:
        """Format number cleanly."""
        try:
            f = float(v)
            return f"{f:,.1f}" if f != int(f) else f"{int(f):,}"
        except (TypeError, ValueError):
            return "—"

    # ── Aggregates ──────────────────────────────────────────────────────────
    total            = len(data)
    powered_on       = sum(1 for v in data if str(v.get("powerState","")).upper() == "POWERED_ON")
    total_actions    = sum(v.get("actionCount", 0) for v in data)
    total_cost_mo    = sum(v.get("vmCostPerMonth") or 0 for v in data)
    total_savings_mo = sum(v.get("totalActionSavingsPerMonth") or 0 for v in data)

    # ── Chart data ───────────────────────────────────────────────────────────
    import json as _json

    vm_names_js   = _json.dumps([v.get("name","") for v in data])
    cpu_util_js   = _json.dumps([round(float(v["cpuUtilizationPct"]),1)
                                  if v.get("cpuUtilizationPct") is not None else 0 for v in data])
    mem_util_js   = _json.dumps([round(float(v["memUtilizationPct"]),1)
                                  if v.get("memUtilizationPct") is not None else 0 for v in data])
    cpu_peak_js   = _json.dumps([round(float(v["cpuPeakUtilPct"]),1)
                                  if v.get("cpuPeakUtilPct") is not None else 0 for v in data])
    mem_peak_js   = _json.dumps([round(float(v["memPeakUtilPct"]),1)
                                  if v.get("memPeakUtilPct") is not None else 0 for v in data])
    cost_js       = _json.dumps([round(float(v["vmCostPerMonth"]),2)
                                  if v.get("vmCostPerMonth") is not None else 0 for v in data])
    savings_js    = _json.dumps([round(float(v["totalActionSavingsPerMonth"]),2)
                                  if v.get("totalActionSavingsPerMonth") is not None else 0 for v in data])
    iops_used_js  = _json.dumps([round(float(v["iopsUsed"]),1)
                                  if v.get("iopsUsed") is not None else 0 for v in data])
    net_util_js   = _json.dumps([round(float(v["netThroughputUtilPct"]),1)
                                  if v.get("netThroughputUtilPct") is not None else 0 for v in data])

    # Disk used % per disk
    disk_labels_js = _json.dumps([
        f"{v.get('name','')}:{d.get('name','')}"
        for v in data for d in v.get("disks", [])
    ])
    disk_iops_cap_js = _json.dumps([
        round(float(d.get("iopsCap") or 0), 0)
        for v in data for d in v.get("disks", [])
    ])
    disk_iops_used_js = _json.dumps([
        round(float(d.get("iopsUsed") or 0), 0)
        for v in data for d in v.get("disks", [])
    ])

    # ── Build table rows ─────────────────────────────────────────────────────
    vm_rows = ""
    for v in data:
        vm_rows += (
            f"<tr>"
            f"<td><strong>{enc(v['name'])}</strong></td>"
            f"<td>{power_badge(v.get('powerState',''))}</td>"
            f"<td>{sev_badge(v.get('severity',''))}</td>"
            f"<td>{badge('STALE','bad') if v.get('staleness')=='STALE' else badge('FRESH','ok')}</td>"
            f"<td>{enc(v['osType'])}</td>"
            f"<td>{enc(v['cloudTier'])}</td>"
            f"<td>{enc(v['region'])}</td>"
            f"<td>{enc(v['resourceGroup'])}</td>"
            f"<td>{enc(v['accountName'])}</td>"
            f"<td>{badge(v.get('billingType',''),'info')}</td>"
            f"<td>{enc(v['riCoveragePct'])}</td>"
            f"<td>{enc(v['uptimePct'])}</td>"
            f"<td>{enc(v['ipAddresses'])}</td>"
            f"<td>{enc(v['actionCount'])}</td>"
            f"</tr>"
        )

    cost_rows = ""
    for v in data:
        rec = v.get("recTierName","")
        cur = v.get("tierName","")
        changed = rec and rec != cur
        cost_rows += (
            f"<tr>"
            f"<td><strong>{enc(v['name'])}</strong></td>"
            f"<td>{usd(v.get('vmCostPerHour'))}</td>"
            f"<td><strong>{usd2(v.get('vmCostPerMonth'))}</strong></td>"
            f"<td>{enc(cur)}</td>"
            f"<td>{enc(v.get('tierFamily'))}</td>"
            f"<td>{'<span class=\"badge badge-warn\">' + enc(rec) + '</span>' if changed else enc(rec) or '—'}</td>"
            f"<td style='color:var(--ok);font-weight:600'>{usd2(v.get('totalActionSavingsPerMonth'))}</td>"
            f"<td>{enc(v.get('tierCpuCapMHz'))}</td>"
            f"<td>{gb(v.get('tierMemCapGB'))}</td>"
            f"<td>{enc(v.get('tierIopsCapacity'))}</td>"
            f"</tr>"
        )

    prov_rows = ""
    for v in data:
        prov_rows += (
            f"<tr>"
            f"<td><strong>{enc(v['name'])}</strong></td>"
            f"<td>{enc(v['numCPUs'])}</td>"
            f"<td>{enc(v.get('physicalCores'))}</td>"
            f"<td>{n(v['cpuCapacityMHz'])}</td>"
            f"<td>{gb(v['memCapacityGB'])}</td>"
            f"<td>{enc(v.get('storageProvisionedGiB'))}</td>"
            f"<td>{n(v.get('iopsCapacity'))}</td>"
            f"<td>{enc(v['numDisks'])}</td>"
            f"<td>{enc(v['numNICs'])}</td>"
            f"</tr>"
        )

    util_rows = ""
    for v in data:
        util_rows += (
            f"<tr>"
            f"<td><strong>{enc(v['name'])}</strong></td>"
            f"<td>{pct_bar(v.get('cpuUtilizationPct'))}</td>"
            f"<td>{pct_bar(v.get('memUtilizationPct'))}</td>"
            f"<td>{pct_bar(v.get('storageAmountUtilPct'))}</td>"
            f"<td>{pct_bar(v.get('storageAccessUtilPct'))}</td>"
            f"<td>{pct_bar(v.get('iopsUtilPct'))}</td>"
            f"<td>{pct_bar(v.get('netThroughputUtilPct'))}</td>"
            f"<td>{pct_bar(v.get('storageLatencyUtilPct'))}</td>"
            f"</tr>"
        )

    peak_rows = ""
    for v in data:
        peak_rows += (
            f"<tr>"
            f"<td><strong>{enc(v['name'])}</strong></td>"
            f"<td>{n(v.get('cpuPeakMHz'))}</td>"
            f"<td>{pct_bar(v.get('cpuPeakUtilPct'))}</td>"
            f"<td>{gb(v.get('memPeakGB'))}</td>"
            f"<td>{pct_bar(v.get('memPeakUtilPct'))}</td>"
            f"<td>{n(v.get('iopsPeak'))}</td>"
            f"<td>{n(v.get('netThroughputPeakKbitps'))}</td>"
            f"<td>{enc(v.get('storageLatencyPeakMs'))}</td>"
            f"</tr>"
        )

    action_rows = ""
    for v in data:
        for a in v.get("actions", []):
            disrupt = str(a.get("disruptiveness",""))
            action_rows += (
                f"<tr>"
                f"<td>{enc(v['name'])}</td>"
                f"<td>{badge(a['actionType'],'info')}</td>"
                f"<td>{enc(a['actionMode'])}</td>"
                f"<td>{enc(a['details'])}</td>"
                f"<td>{enc(a.get('riskSubCategory',''))}</td>"
                f"<td>{sev_badge(a.get('riskSeverity',''))}</td>"
                f"<td>{enc(a['currentEntity'])}</td>"
                f"<td>{badge(a['newEntity'],'warn') if a.get('newEntity') else '—'}</td>"
                f"<td>{'<span class=\"badge badge-bad\">DISRUPTIVE</span>' if disrupt=='DISRUPTIVE' else badge(disrupt,'ok')}</td>"
                f"<td>{enc(a.get('reversibility',''))}</td>"
                f"<td style='color:var(--ok);font-weight:700'>{usd2(a.get('savingsPerMonth'))}/mo</td>"
                f"</tr>"
            )
    if not action_rows:
        action_rows = "<tr><td colspan='11' style='text-align:center;color:var(--muted)'>No pending actions.</td></tr>"

    disk_rows = ""
    for v in data:
        for d in v.get("disks", []):
            cap  = float(d.get("iopsCap") or 0)
            used = float(d.get("iopsUsed") or 0)
            pct  = round(used / cap * 100, 1) if cap > 0 else None
            disk_rows += (
                f"<tr>"
                f"<td>{enc(v['name'])}</td>"
                f"<td>{enc(d.get('name'))}</td>"
                f"<td>{badge(d.get('tier',''),'info')}</td>"
                f"<td>{enc(d.get('storageAmtGiB'))}</td>"
                f"<td>{n(d.get('iopsCap'))}</td>"
                f"<td>{n(d.get('iopsUsed'))}</td>"
                f"<td>{pct_bar(pct)}</td>"
                f"<td>{n(d.get('ioThroughputCapKBps'))}</td>"
                f"<td>{n(d.get('ioThroughputUsedKBps'))}</td>"
                f"</tr>"
            )
    if not disk_rows:
        disk_rows = "<tr><td colspan='9' style='text-align:center;color:var(--muted)'>No disk data.</td></tr>"

    # ── Assemble HTML ────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Turbonomic VM Metrics</title>
<style>
:root{{--bg:#f0f2f5;--surface:#fff;--border:#e2e6ea;--text:#1a1d23;--muted:#6b7280;--accent:#0f62fe;--accent-light:#eef3ff;--ok:#198038;--warn:#f1620a;--bad:#da1e28;--nav-w:220px;--hdr:60px}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--text);font-family:-apple-system,"Segoe UI",system-ui,sans-serif;font-size:13px;line-height:1.5}}
.topbar{{position:fixed;top:0;left:0;right:0;height:var(--hdr);background:#0d1117;color:#fff;display:flex;align-items:center;padding:0 24px;gap:14px;z-index:200;box-shadow:0 2px 8px rgba(0,0,0,.4)}}
.topbar h1{{font-size:16px;font-weight:700;white-space:nowrap}}
.topbar .meta{{font-size:11px;color:#8b949e;margin-left:auto;white-space:nowrap}}
.sidebar{{position:fixed;top:var(--hdr);left:0;width:var(--nav-w);bottom:0;background:#fff;border-right:1px solid var(--border);overflow-y:auto;z-index:100;padding:14px 0}}
.nav-group{{padding:6px 16px 2px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--muted)}}
.nav-item{{display:block;padding:7px 20px;color:#374151;text-decoration:none;font-size:12.5px;border-left:3px solid transparent;transition:all .15s}}
.nav-item:hover,.nav-item.active{{background:var(--accent-light);border-left-color:var(--accent);color:var(--accent)}}
.main{{margin-left:var(--nav-w);margin-top:var(--hdr);padding:22px 26px;min-height:calc(100vh - var(--hdr))}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:22px}}
.kpi{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px}}
.kpi-val{{font-size:28px;font-weight:800;line-height:1}}
.kpi-label{{font-size:11px;color:var(--muted);margin-top:4px}}
.kpi.ok .kpi-val{{color:var(--ok)}}.kpi.bad .kpi-val{{color:var(--bad)}}.kpi.accent .kpi-val{{color:var(--accent)}}
.charts-row{{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px;margin-bottom:20px}}
.chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px}}
.chart-card h3{{font-size:11.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px}}
.section{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px 20px;margin-bottom:20px;scroll-margin-top:calc(var(--hdr)+14px)}}
.section-hdr{{display:flex;align-items:center;gap:10px;margin-bottom:12px}}
.section-hdr h2{{font-size:13.5px;font-weight:700;display:flex;align-items:center;gap:7px}}
.section-hdr h2 .ico{{width:20px;height:20px;border-radius:5px;background:var(--accent);display:flex;align-items:center;justify-content:center;color:#fff;font-size:10px;font-weight:800;flex-shrink:0}}
.search-bar{{margin-bottom:10px}}
.search-bar input{{padding:7px 11px;border:1px solid var(--border);border-radius:6px;font-size:12px;outline:none;width:min(340px,100%)}}
.search-bar input:focus{{border-color:var(--accent);box-shadow:0 0 0 2px rgba(15,98,254,.15)}}
.tablewrap{{overflow-x:auto;border-radius:6px;border:1px solid var(--border)}}
table{{border-collapse:collapse;width:100%;font-size:12.5px}}
thead th{{background:#f8f9fb;text-align:left;padding:8px 11px;border-bottom:2px solid var(--border);white-space:nowrap;font-weight:600;color:#374151;font-size:11.5px}}
tbody td{{padding:7px 11px;border-bottom:1px solid #f0f2f5;vertical-align:middle}}
tbody tr:last-child td{{border-bottom:none}}
tbody tr:hover td{{background:#f8faff}}
.badge{{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600;white-space:nowrap}}
.badge-ok{{background:#dcf5e6;color:#0a6640}}.badge-warn{{background:#fff3cd;color:#856404}}
.badge-bad{{background:#fde8e8;color:#9b1c1c}}.badge-info{{background:#dbeafe;color:#1e40af}}
.badge-muted{{background:#f3f4f6;color:#4b5563}}
.pbar-wrap{{display:flex;align-items:center;gap:6px}}
.pbar-track{{background:#e5e7eb;border-radius:4px;height:7px;width:80px;flex-shrink:0;overflow:hidden}}
.pbar-fill{{height:7px;border-radius:4px}}
.pbar-label{{font-size:11px;color:var(--muted);min-width:32px}}
.muted{{color:var(--muted)}}
footer{{text-align:center;font-size:11px;color:var(--muted);padding:18px;border-top:1px solid var(--border);margin-top:6px}}
@media(max-width:800px){{.sidebar{{display:none}}.main{{margin-left:0}}}}
</style>
</head>
<body>

<div class="topbar">
  <h1>&#9711; Turbonomic VM Metrics</h1>
  <span class="meta">Generated: {enc(generated_at)} &nbsp;|&nbsp; {total} VMs</span>
</div>

<nav class="sidebar">
  <div class="nav-group">Overview</div>
  <a class="nav-item active" href="#summary">&#128200; Summary &amp; Charts</a>
  <div class="nav-group">VM Details</div>
  <a class="nav-item" href="#s-overview">&#128187; VM Overview</a>
  <a class="nav-item" href="#s-cost">&#128181; Cost &amp; Sizing</a>
  <a class="nav-item" href="#s-prov">&#9964; Provisioned</a>
  <a class="nav-item" href="#s-util">&#128209; Avg Utilization</a>
  <a class="nav-item" href="#s-peak">&#128200; Peak Values</a>
  <div class="nav-group">Actions &amp; Disks</div>
  <a class="nav-item" href="#s-actions">&#9889; Pending Actions</a>
  <a class="nav-item" href="#s-disks">&#128190; Per-Disk Detail</a>
</nav>

<main class="main">

<!-- KPI -->
<div class="kpi-grid">
  <div class="kpi accent"><div class="kpi-val">{total}</div><div class="kpi-label">VMs Found</div></div>
  <div class="kpi ok"><div class="kpi-val">{powered_on}</div><div class="kpi-label">Powered On</div></div>
  <div class="kpi"><div class="kpi-val">{total - powered_on}</div><div class="kpi-label">Off / Unknown</div></div>
  <div class="kpi{'bad' if total_actions > 0 else ''}"><div class="kpi-val">{total_actions}</div><div class="kpi-label">Pending Actions</div></div>
  <div class="kpi"><div class="kpi-val" style="font-size:20px">${total_cost_mo:.2f}</div><div class="kpi-label">Total Cost / Month</div></div>
  <div class="kpi ok"><div class="kpi-val" style="font-size:20px">${total_savings_mo:.2f}</div><div class="kpi-label">Potential Savings / Month</div></div>
</div>

<!-- CHARTS -->
<div id="summary">
<div class="charts-row">
  <div class="chart-card"><h3>CPU Avg vs Peak Utilization (%)</h3><canvas id="cCpuUtil" height="200"></canvas></div>
  <div class="chart-card"><h3>Memory Avg vs Peak Utilization (%)</h3><canvas id="cMemUtil" height="200"></canvas></div>
  <div class="chart-card"><h3>VM Cost &amp; Savings ($/month)</h3><canvas id="cCost" height="200"></canvas></div>
</div>
<div class="charts-row">
  <div class="chart-card"><h3>IOPS Used per VM</h3><canvas id="cIops" height="200"></canvas></div>
  <div class="chart-card"><h3>Network Throughput Utilization (%)</h3><canvas id="cNet" height="200"></canvas></div>
</div>
<div class="charts-row">
  <div class="chart-card" style="grid-column:1/-1"><h3>Disk IOPS — Capacity vs Used</h3><canvas id="cDiskIops" height="160"></canvas></div>
</div>
</div>

<!-- VM OVERVIEW -->
<div class="section" id="s-overview">
<div class="section-hdr"><h2><span class="ico">V</span>VM Overview</h2></div>
<div class="search-bar"><input id="fi1" placeholder="&#128269; Filter..." oninput="ft('ti1','fi1')"></div>
<div class="tablewrap"><table id="ti1">
<thead><tr>
  <th>Name</th><th>Power</th><th>Severity</th><th>Staleness</th>
  <th>OS</th><th>Cloud Tier</th><th>Region</th><th>Resource Group</th>
  <th>Account</th><th>Billing</th><th>RI Coverage %</th><th>Uptime %</th>
  <th>IP Addresses</th><th>Actions</th>
</tr></thead>
<tbody>{vm_rows or '<tr><td colspan=14>No data.</td></tr>'}</tbody>
</table></div>
</div>

<!-- COST -->
<div class="section" id="s-cost">
<div class="section-hdr"><h2><span class="ico">$</span>Cost &amp; Instance Sizing</h2></div>
<div class="tablewrap"><table>
<thead><tr>
  <th>Name</th><th>Cost ($/h)</th><th>Cost ($/mo)</th>
  <th>Current Tier</th><th>Tier Family</th><th>Recommended Tier</th>
  <th>Savings ($/mo)</th>
  <th>Tier CPU (MHz)</th><th>Tier Mem (GB)</th><th>Tier IOPS</th>
</tr></thead>
<tbody>{cost_rows or '<tr><td colspan=10>No data.</td></tr>'}</tbody>
</table></div>
</div>

<!-- PROVISIONED -->
<div class="section" id="s-prov">
<div class="section-hdr"><h2><span class="ico">P</span>Provisioned Resources</h2></div>
<div class="tablewrap"><table>
<thead><tr>
  <th>Name</th><th>vCPUs</th><th>Physical Cores</th><th>CPU Cap (MHz)</th>
  <th>Mem Cap (GB)</th><th>Storage Prov (GiB)</th><th>IOPS Cap</th>
  <th>Disks</th><th>NICs</th>
</tr></thead>
<tbody>{prov_rows or '<tr><td colspan=9>No data.</td></tr>'}</tbody>
</table></div>
</div>

<!-- UTILIZATION -->
<div class="section" id="s-util">
<div class="section-hdr"><h2><span class="ico">U</span>Average Utilization %</h2></div>
<div class="search-bar"><input id="fi4" placeholder="&#128269; Filter..." oninput="ft('ti4','fi4')"></div>
<div class="tablewrap"><table id="ti4">
<thead><tr>
  <th>Name</th>
  <th>CPU (vCPU) %</th><th>Memory %</th>
  <th>Storage Amt %</th><th>Storage Access %</th><th>IOPS %</th>
  <th>Net Throughput %</th><th>Storage Latency %</th>
</tr></thead>
<tbody>{util_rows or '<tr><td colspan=8>No data.</td></tr>'}</tbody>
</table></div>
</div>

<!-- PEAK -->
<div class="section" id="s-peak">
<div class="section-hdr"><h2><span class="ico">^</span>Peak Values (7-day)</h2></div>
<div class="tablewrap"><table>
<thead><tr>
  <th>Name</th>
  <th>CPU Peak (MHz)</th><th>CPU Peak %</th>
  <th>Mem Peak (GB)</th><th>Mem Peak %</th>
  <th>IOPS Peak</th><th>Net Peak (Kbit/s)</th><th>Latency Peak (ms)</th>
</tr></thead>
<tbody>{peak_rows or '<tr><td colspan=8>No data.</td></tr>'}</tbody>
</table></div>
</div>

<!-- ACTIONS -->
<div class="section" id="s-actions">
<div class="section-hdr"><h2><span class="ico">!</span>Pending Actions</h2></div>
<div class="tablewrap"><table>
<thead><tr>
  <th>VM</th><th>Action Type</th><th>Mode</th><th>Details</th>
  <th>Risk Category</th><th>Risk Severity</th>
  <th>Current Instance</th><th>Recommended Instance</th>
  <th>Disruptive?</th><th>Reversible?</th><th>Savings ($/mo)</th>
</tr></thead>
<tbody>{action_rows}</tbody>
</table></div>
</div>

<!-- DISKS -->
<div class="section" id="s-disks">
<div class="section-hdr"><h2><span class="ico">D</span>Per-Disk Detail</h2></div>
<div class="tablewrap"><table>
<thead><tr>
  <th>VM</th><th>Disk</th><th>Tier</th>
  <th>Size (GiB)</th><th>IOPS Cap</th><th>IOPS Used</th><th>IOPS Util %</th>
  <th>IO Throughput Cap (KB/s)</th><th>IO Throughput Used (KB/s)</th>
</tr></thead>
<tbody>{disk_rows}</tbody>
</table></div>
</div>

</main>
<footer>Turbonomic VM Metrics Collector &nbsp;|&nbsp; Made with IBM Bob</footer>

<script>
/* ─── Table filter ─── */
function ft(tId,iId){{
  var f=document.getElementById(iId).value.toLowerCase();
  var rows=document.getElementById(tId).getElementsByTagName('tr');
  for(var i=1;i<rows.length;i++)
    rows[i].style.display=rows[i].innerText.toLowerCase().indexOf(f)>=0?'':'none';
}}
/* ─── Nav highlight ─── */
(function(){{
  var items=document.querySelectorAll('.nav-item[href^="#"]');
  var targets=[].map.call(items,function(a){{return document.querySelector(a.getAttribute('href'))}});
  function upd(){{var y=window.scrollY+80,cur=-1;targets.forEach(function(t,i){{if(t&&t.offsetTop<=y)cur=i}});items.forEach(function(a,i){{a.classList.toggle('active',i===cur)}})}}
  window.addEventListener('scroll',upd,{{passive:true}});upd();
}})();
/* ─── SVG chart engine ─── */
(function(){{
  var C=['#0f62fe','#198038','#f1620a','#8b5cf6','#da1e28','#06b6d4','#f59e0b'];
  function svg(id,w,h){{var c=document.getElementById(id);if(!c)return null;var s=document.createElementNS('http://www.w3.org/2000/svg','svg');s.setAttribute('viewBox','0 0 '+w+' '+h);s.setAttribute('width','100%');s.setAttribute('font-family','inherit');s.setAttribute('font-size','11');c.parentNode.replaceChild(s,c);return s;}}
  function el(tag,a,p){{var e=document.createElementNS('http://www.w3.org/2000/svg',tag);for(var k in a)e.setAttribute(k,a[k]);if(p)p.appendChild(e);return e;}}
  function txt(s,x,y,txt,attrs){{var t=el('text',Object.assign({{x:x,y:y,fill:'#374151'}},attrs||{{}}));t.textContent=txt;s.appendChild(t);}}

  /* grouped horizontal bar: labels[], series=[{{name,color,vals}}] */
  function hbarGroup(id,labels,series){{
    var n=labels.length,ns=series.length,bH=14,gap=4,gGap=12,padL=180,padR=50,padT=26,padB=8;
    var groupH=ns*(bH+gap)-gap,W=560,H=padT+n*(groupH+gGap)-gGap+padB;
    var s=svg(id,W,H);if(!s)return;
    var allV=series.reduce(function(a,b){{return a.concat(b.vals)}},[] );
    var maxV=Math.max.apply(null,allV)||1;
    /* legend */
    var lx=padL;
    series.forEach(function(sr){{el('rect',{{x:lx,y:8,width:11,height:11,rx:2,fill:sr.color}},s);el('text',{{x:lx+14,y:18,fill:'#374151'}},s).textContent=sr.name;lx+=sr.name.length*6.5+28;}});
    labels.forEach(function(lbl,i){{
      var gy=padT+i*(groupH+gGap);
      el('text',{{x:padL-6,y:gy+groupH/2+4,'text-anchor':'end',fill:'#374151'}},s).textContent=lbl;
      series.forEach(function(sr,si){{
        var y=gy+si*(bH+gap);
        var bw=Math.max(2,(sr.vals[i]/maxV)*(W-padL-padR));
        el('rect',{{x:padL,y:y,width:W-padL-padR,height:bH,rx:3,fill:'#f0f2f5'}},s);
        el('rect',{{x:padL,y:y,width:bw,height:bH,rx:3,fill:sr.color}},s);
        el('text',{{x:padL+bw+4,y:y+bH-3,fill:'#374151'}},s).textContent=sr.vals[i];
      }});
    }});
  }}

  /* simple horizontal bar */
  function hbar(id,labels,vals,color){{
    var n=labels.length,bH=24,gap=7,padL=160,padR=60,padT=8,padB=8;
    var W=500,H=padT+n*(bH+gap)-gap+padB;
    var s=svg(id,W,H);if(!s)return;
    var maxV=Math.max.apply(null,vals)||1;
    labels.forEach(function(lbl,i){{
      var y=padT+i*(bH+gap);
      var bw=Math.max(2,(vals[i]/maxV)*(W-padL-padR));
      el('text',{{x:padL-6,y:y+bH/2+4,'text-anchor':'end',fill:'#374151'}},s).textContent=lbl;
      el('rect',{{x:padL,y:y,width:W-padL-padR,height:bH,rx:4,fill:'#f0f2f5'}},s);
      el('rect',{{x:padL,y:y,width:bw,height:bH,rx:4,fill:color||C[0]}},s);
      el('text',{{x:padL+bw+5,y:y+bH/2+4,fill:'#374151'}},s).textContent=vals[i];
    }});
  }}

  /* donut */
  function donut(id,labels,vals){{
    var W=420,H=200,cx=100,cy=100,r=85,ir=52;
    var s=svg(id,W,H);if(!s)return;
    var tot=vals.reduce(function(a,b){{return a+b}},0)||1,angle=-Math.PI/2;
    vals.forEach(function(v,i){{
      if(!v)return;
      var sw=2*Math.PI*(v/tot);
      var x1=cx+r*Math.cos(angle),y1=cy+r*Math.sin(angle),x2=cx+r*Math.cos(angle+sw),y2=cy+r*Math.sin(angle+sw);
      var xi1=cx+ir*Math.cos(angle),yi1=cy+ir*Math.sin(angle),xi2=cx+ir*Math.cos(angle+sw),yi2=cy+ir*Math.sin(angle+sw);
      var lg=sw>Math.PI?1:0;
      el('path',{{d:'M'+xi1+' '+yi1+' L'+x1+' '+y1+' A'+r+' '+r+' 0 '+lg+' 1 '+x2+' '+y2+' L'+xi2+' '+yi2+' A'+ir+' '+ir+' 0 '+lg+' 0 '+xi1+' '+yi1+'Z',fill:C[i%C.length]}},s);
      angle+=sw;
    }});
    var ly=18;
    labels.forEach(function(lbl,i){{
      if(!vals[i])return;
      el('rect',{{x:210,y:ly-11,width:12,height:12,rx:2,fill:C[i%C.length]}},s);
      el('text',{{x:226,y:ly,fill:'#374151'}},s).textContent=lbl+' ('+vals[i]+')';
      ly+=20;
    }});
  }}

  var VMS   = {vm_names_js};
  var CPU_A = {cpu_util_js};
  var CPU_P = {cpu_peak_js};
  var MEM_A = {mem_util_js};
  var MEM_P = {mem_peak_js};
  var COST  = {cost_js};
  var SAV   = {savings_js};
  var IOPS  = {iops_used_js};
  var NET   = {net_util_js};
  var DL    = {disk_labels_js};
  var DIC   = {disk_iops_cap_js};
  var DIU   = {disk_iops_used_js};
  hbarGroup('cCpuUtil',VMS,[{{name:'Avg %',color:'#0f62fe',vals:CPU_A}},{{name:'Peak %',color:'#da1e28',vals:CPU_P}}]);
  hbarGroup('cMemUtil',VMS,[{{name:'Avg %',color:'#198038',vals:MEM_A}},{{name:'Peak %',color:'#f1620a',vals:MEM_P}}]);
  hbarGroup('cCost',VMS,[{{name:'Cost $/mo',color:'#0f62fe',vals:COST}},{{name:'Savings $/mo',color:'#198038',vals:SAV}}]);
  hbar('cIops',VMS,IOPS,'#8b5cf6');
  hbar('cNet',VMS,NET,'#06b6d4');
  hbarGroup('cDiskIops',DL,[{{name:'IOPS Cap',color:'#c8d6e5',vals:DIC}},{{name:'IOPS Used',color:'#0f62fe',vals:DIU}}]);
}})();
</script>
</body>
</html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    log.info("HTML written: %s", path)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    # Default to config/config.json
    default_config = (
        Path(__file__).parent.parent / "config" / "config.json"
        if (Path(__file__).parent.parent / "config" / "config.json").exists()
        else Path(__file__).parent / "config.json"
    )

    parser = argparse.ArgumentParser(description="Turbonomic VM Metrics Collector")
    parser.add_argument("--config",  default=str(default_config))
    parser.add_argument("--vm",      nargs="*", help="VM name(s) to query")
    parser.add_argument("--days",    type=int, default=1,
                        help="Historical period in days for peak data (default: 1)")
    parser.add_argument("--output",  default=None,
                        help="Output directory for JSON/CSV/HTML files. "
                             "Overrides the 'output.directory' value in config.json.")
    parser.add_argument("--skip-ssl-verify", action="store_true")
    args = parser.parse_args()

    config_path = Path(args.config)
    if not config_path.is_absolute() and not config_path.exists():
        alt_path = Path(__file__).parent.parent / args.config
        if alt_path.exists():
            config_path = alt_path

    if not config_path.exists():
        log.error("Config file not found: %s", config_path)
        sys.exit(1)

    with open(config_path, encoding="utf-8") as f:
        cfg = json.load(f)

    turbo_cfg  = cfg.get("turbonomic", {})
    output_cfg = cfg.get("output", {})

    # Extract VM names from command line or config (config/config.json targets or servers_csv)
    vm_names = args.vm if args.vm else extract_vm_names_from_config(cfg, config_path.parent)
    if not vm_names:
        log.error("No VM names found in config or args.")
        sys.exit(1)

    log.info("Target VMs to query: %s", vm_names)

    # --output flag takes priority; then config output.directory; then ./output beside config
    if args.output:
        out_dir = Path(args.output).resolve()
    else:
        out_dir = (config_path.parent / output_cfg.get("directory", "./output")).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    run_stamp    = datetime.now().strftime("%Y%m%d-%H%M%S")
    generated_at = datetime.now(timezone.utc).isoformat()

    # Turbonomic connection settings (supports both url/host/api_base_url and username/admin_username)
    base_url = (
        turbo_cfg.get("url")
        or turbo_cfg.get("api_base_url")
        or (f"https://{turbo_cfg.get('host')}" if turbo_cfg.get("host") else None)
    )
    if base_url and "/api/v3" in base_url:
        base_url = base_url.split("/api/v3")[0]

    username = turbo_cfg.get("username") or turbo_cfg.get("admin_username")
    password = turbo_cfg.get("password") or turbo_cfg.get("admin_password")

    if not base_url or not username or not password:
        log.error("Turbonomic credentials incomplete in config (need url/host, username/admin_username, password/admin_password)")
        sys.exit(1)

    verify_ssl = False if args.skip_ssl_verify else turbo_cfg.get("verify_ssl", False)

    client = TurbonomicClient(
        url        = base_url,
        username   = username,
        password   = password,
        verify_ssl = verify_ssl,
        timeout    = turbo_cfg.get("timeout_seconds", 60),
    )

    try:
        client.authenticate()
    except Exception as exc:
        log.error("Authentication failed: %s", exc)
        sys.exit(1)

    entities = client.search_vms(vm_names)
    if not entities:
        log.warning("No matching VMs found for: %s", vm_names)

    all_metrics: list[dict] = []
    for entity in entities:
        try:
            all_metrics.append(client.collect_vm_metrics(entity, days=args.days))
        except Exception as exc:
            log.error("Failed for %s: %s", entity.get("displayName"), exc)

    log.info("Collected metrics for %d / %d VMs", len(all_metrics), len(entities))

    found_lower = {m["name"].lower() for m in all_metrics}
    for n in vm_names:
        if n.lower() not in found_lower:
            log.warning("VM not found in Turbonomic: %s", n)

    formats = output_cfg.get("formats", ["json", "csv", "html"])

    if "json" in formats:
        write_json(all_metrics, str(out_dir / f"turbo-vm-metrics-{run_stamp}.json"))
    if "csv" in formats:
        write_csv(all_metrics, str(out_dir / f"turbo-vm-metrics-{run_stamp}.csv"))
    if "html" in formats:
        write_html(all_metrics, str(out_dir / f"turbo-vm-metrics-{run_stamp}.html"),
                   generated_at)

    print("\n" + "=" * 60)
    print("TURBONOMIC VM METRICS — COMPLETE")
    print("=" * 60)
    print(f"VMs requested     : {len(vm_names)}")
    print(f"VMs found         : {len(entities)}")
    print(f"Metrics collected : {len(all_metrics)}")
    print(f"Output directory  : {out_dir}")
    print("=" * 60)


if __name__ == "__main__":
    main()
