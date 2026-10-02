#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Turbonomic Target Registration & Discovery Script
==================================================
Creates WMI targets for Windows VMs via the Turbonomic REST API,
builds the shared static scope group, and triggers target rediscovery.

Usage:
    python register_turbo_targets.py [--config ../config/config.json]
    py.exe register_turbo_targets.py [--config ..\config\config.json]  (Windows)

Docs:
    https://www.ibm.com/docs/en/tarm/8.21.1?topic=targets-wmi
    https://www.ibm.com/docs/en/tarm/8.21.1?topic=endpoints-targets-endpoint
"""

import argparse
import json
import logging
import sys
import time
import urllib.parse
from pathlib import Path

import requests
import urllib3

# Suppress InsecureRequestWarning for self-signed certs on Turbonomic appliances
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

def setup_logging(level: str, log_file: str) -> logging.Logger:
    """Configure root logger to write DEBUG-level output to both the
    console and a rotating log file.  Full DEBUG output is always written
    to the file; the console honours the requested *level*.

    *log_file* should be an absolute path so the log lands next to the
    config file regardless of the working directory (important on Windows
    where CWD is unpredictable when scripts are invoked via py.exe).
    """
    numeric_level = getattr(logging, level.upper(), logging.DEBUG)

    fmt = "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s"
    datefmt = "%Y-%m-%dT%H:%M:%S"

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)  # capture everything at root

    # Console handler (respects requested level)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(numeric_level)
    ch.setFormatter(logging.Formatter(fmt, datefmt))
    root.addHandler(ch)

    # File handler (always DEBUG for full discovery/target logs)
    # Ensure the parent directory exists before opening (cross-platform safety)
    log_path = Path(log_file)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    fh = logging.FileHandler(str(log_path), encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(fmt, datefmt))
    root.addHandler(fh)

    return logging.getLogger("turbo.wmi")


# ---------------------------------------------------------------------------
# Turbonomic API client
# ---------------------------------------------------------------------------

class TurbonomicClient:
    """Minimal REST client that handles session authentication and common
    target operations against the Turbonomic v3 API."""

    def __init__(self, base_url: str, verify_ssl: bool = False):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.verify = verify_ssl
        self.session.headers.update({"Accept": "application/json",
                                     "Content-Type": "application/json"})
        self.log = logging.getLogger("turbo.wmi.client")

    # ------------------------------------------------------------------
    # Authentication
    # ------------------------------------------------------------------

    def login(self, username: str, password: str) -> None:
        """Authenticate and store the session cookie for subsequent calls.

        The Turbonomic API uses form-encoded credentials (x-www-form-urlencoded)
        and returns a JSESSIONID cookie.  Ref:
        https://www.ibm.com/docs/en/tarm/8.21.1?topic=api-authenticating
        """
        url = f"{self.base_url}/login"
        payload = urllib.parse.urlencode({"username": username,
                                          "password": password})
        headers = {"Content-Type": "application/x-www-form-urlencoded"}

        self.log.debug("POST %s  (authenticating user=%s)", url, username)
        resp = self.session.post(url, data=payload, headers=headers)
        self.log.debug("Login response: HTTP %s", resp.status_code)
        self.log.debug("Login response headers:\n%s",
                       json.dumps(dict(resp.headers), indent=2))

        if resp.status_code not in (200, 204):
            self.log.error("Authentication failed: HTTP %s\n%s",
                           resp.status_code, resp.text)
            raise RuntimeError(
                f"Authentication failed (HTTP {resp.status_code}): {resp.text}"
            )

        self.log.info("Authentication successful for user '%s'", username)

    # ------------------------------------------------------------------
    # Target management
    # ------------------------------------------------------------------

    def get_target_specs(self) -> list:
        """Fetch the probe registry — lists all probe types and their
        required inputFields.  Useful for verifying the WMI field names
        before attempting to add the target."""
        url = f"{self.base_url}/targets/specs"
        self.log.debug("GET %s  (fetching target specs)", url)
        resp = self.session.get(url)
        self._raise_for_status(resp, "get target specs")
        specs = resp.json()
        self.log.debug("Received %d probe spec entries", len(specs))
        return specs

    def find_wmi_spec(self) -> dict | None:
        """Return the WMI probe spec from /targets/specs, or None."""
        specs = self.get_target_specs()
        for entry in specs:
            if isinstance(entry, dict) and entry.get("type", "").upper() == "WMI":
                self.log.debug("WMI probe spec found:\n%s",
                               json.dumps(entry, indent=2))
                return entry
        self.log.warning("WMI probe spec not found in /targets/specs response")
        return None

    def search_entity_by_ip(self, ip_address: str, vm_name: str = "") -> str | None:
        """Search for a VirtualMachine entity by VM name or IP and return its UUID.

        Prioritizes searching by vm_name (Turbonomic VM displayName / guest name),
        falling back to IP address.
        """
        url = f"{self.base_url}/search"

        # 1. Primary search by VM name (exact or prefix/substring via query)
        if vm_name:
            params = {"q": vm_name, "types": "VirtualMachine"}
            self.log.debug("GET %s  params=%s  (searching VM by vm_name)", url, params)
            resp = self.session.get(url, params=params)
            if resp.ok:
                results = resp.json()
                if results:
                    # Look for exact displayName match first, otherwise take first result
                    matched = next((e for e in results if e.get("displayName", "").strip().lower() == vm_name.strip().lower()), results[0])
                    uuid = matched.get("uuid")
                    display = matched.get("displayName", "")
                    self.log.info(
                        "Found VM entity: displayName='%s' uuid=%s (matched vm_name '%s')",
                        display, uuid, vm_name
                    )
                    return uuid

            # Search with criteria for vm_name
            for filter_type in ["vmsByName", "vmsByGuestName"]:
                body = {
                    "criteriaList": [
                        {"expType": "EQ", "expVal": vm_name, "filterType": filter_type, "caseSensitive": False}
                    ],
                    "types": ["VirtualMachine"]
                }
                resp = self.session.post(url, data=json.dumps(body))
                if resp.ok:
                    results = resp.json()
                    if results:
                        entity = results[0]
                        uuid = entity.get("uuid")
                        display = entity.get("displayName", "")
                        self.log.info(
                            "Found VM entity via POST /search criteria (%s='%s'): displayName='%s' uuid=%s",
                            filter_type, vm_name, display, uuid
                        )
                        return uuid

        # 2. Search by IP address if provided
        if ip_address:
            params = {"q": ip_address, "types": "VirtualMachine"}
            self.log.debug("GET %s  params=%s  (searching VM by IP)", url, params)
            resp = self.session.get(url, params=params)
            if resp.ok:
                results = resp.json()
                if results:
                    entity = results[0]
                    uuid = entity.get("uuid")
                    display = entity.get("displayName", "")
                    self.log.info(
                        "Found VM entity: displayName='%s' uuid=%s (matched IP %s)",
                        display, uuid, ip_address
                    )
                    return uuid

            for filter_type in ["vmsByGuestName", "vmsByName"]:
                body = {
                    "criteriaList": [
                        {"expType": "EQ", "expVal": ip_address, "filterType": filter_type, "caseSensitive": False}
                    ],
                    "types": ["VirtualMachine"]
                }
                resp = self.session.post(url, data=json.dumps(body))
                if resp.ok:
                    results = resp.json()
                    if results:
                        entity = results[0]
                        uuid = entity.get("uuid")
                        display = entity.get("displayName", "")
                        self.log.info(
                            "Found VM entity via POST /search criteria (%s='%s'): displayName='%s' uuid=%s",
                            filter_type, ip_address, display, uuid
                        )
                        return uuid

        self.log.warning("No VirtualMachine entity found in inventory for vm_name='%s'%s",
                         vm_name, f" or IP '{ip_address}'" if ip_address else "")
        return None

    # Valid filterTypes for VirtualMachine groups (from Turbonomic API docs):
    # vmsByName, vmsByPMName, vmsByGuestName, vmsByAltName, vmsByTag, vmsByState,
    # vmsByClusterName, vmsByDC, vmsByVDC, vmsByNumCPUs, vmsByMem, vmsByStorage,
    # vmsByNetwork, vmsByApplication, vmsByDatabaseServer, vmsByBusinessAccountUuid,
    # vmsByResourceGroupUuid
    # NOTE: 'vmsByIp' does NOT exist — IPs are stored as alternate names.
    _IP_FILTER_CANDIDATES = ["vmsByAltName", "vmsByGuestName", "vmsByName"]

    def delete_group(self, group_uuid: str) -> None:
        """DELETE /groups/{uuid} — remove an existing group so it can be recreated cleanly."""
        url = f"{self.base_url}/groups/{group_uuid}"
        self.log.debug("DELETE %s (deleting old scope group)", url)
        resp = self.session.delete(url)
        if resp.ok or resp.status_code == 404:
            self.log.info("Deleted existing group uuid=%s", group_uuid)
        else:
            self.log.warning("Could not delete group uuid=%s: HTTP %s %s", group_uuid, resp.status_code, resp.text)

    def find_group_by_name(self, group_name: str) -> dict | None:
        """Find an existing group by exact display name."""
        url = f"{self.base_url}/groups"
        resp = self.session.get(url, params={"q": group_name})
        if resp.ok:
            for g in resp.json():
                if g.get("displayName", "").strip() == group_name.strip():
                    return g
        return None

    def create_static_vm_scope_group(self, vm_uuids: list[str], group_name: str) -> str:
        """Create a static VM group containing strictly the specified VM UUIDs.

        If a group with the same name already exists, deletes and recreates it to ensure
        it is purely Static with exactly the required members.
        """
        url = f"{self.base_url}/groups"

        existing = self.find_group_by_name(group_name)
        if existing:
            self.log.info("Existing scope group '%s' found (uuid=%s) — deleting before creating fresh static group",
                          group_name, existing["uuid"])
            self.delete_group(existing["uuid"])

        dto = {
            "isStatic": True,
            "displayName": group_name,
            "memberUuidList": vm_uuids,
            "groupType": "VirtualMachine"
        }
        self.log.debug("POST %s (creating static VM scope group for %d member(s))", url, len(vm_uuids))
        resp = self.session.post(url, data=json.dumps(dto))
        self.log.debug("Create static group response: HTTP %s\n%s", resp.status_code, resp.text)
        if resp.ok:
            created = resp.json()
            uuid = created["uuid"]
            self.log.info("Created static scope group '%s' (Type=Static, Members=%d) uuid=%s",
                          group_name, len(vm_uuids), uuid)
            return uuid

        self.log.warning("Failed to create static group '%s': HTTP %s %s", group_name, resp.status_code, resp.text)
        return ""

    def create_ip_scope_group(self, ip_addresses: list[str] | str,
                              group_name: str) -> str:
        """Create a dynamic VM group filtered strictly by IP addresses and return its UUID.

        Turbonomic stores IP addresses as alternate names, so this tries
        filterType candidates in order until the API accepts one:
          1. vmsByAltName  — IP stored as alternate name (exact EQ match for VMs)
          2. vmsByGuestName — guest OS name matching IP exactly
          3. vmsByName     — display name fallback

        The group is reused if it already exists (same display name).
        """
        url = f"{self.base_url}/groups"

        # Reuse if already exists
        resp = self.session.get(url, params={"q": group_name})
        if resp.ok:
            for g in resp.json():
                if g.get("displayName", "").strip() == group_name:
                    self.log.info(
                        "Reusing existing scope group '%s'  uuid=%s",
                        group_name, g["uuid"]
                    )
                    return g["uuid"]

        ips = [ip_addresses] if isinstance(ip_addresses, str) else ip_addresses

        last_error = ""
        for filter_type in self._IP_FILTER_CANDIDATES:
            criteria_list = [
                {
                    "expType": "EQ",
                    "expVal": ip,
                    "filterType": filter_type,
                    "caseSensitive": False
                }
                for ip in ips
            ]
            dto = {
                "isStatic": False,
                "displayName": group_name,
                "memberUuidList": [],
                "criteriaList": criteria_list,
                "groupType": "VirtualMachine"
            }
            self.log.debug(
                "POST %s  (creating IP scope group, filterType=%s, ips=%s)",
                url, filter_type, ips
            )
            self.log.debug("Group DTO:\n%s", json.dumps(dto, indent=2))
            resp = self.session.post(url, data=json.dumps(dto))
            self.log.debug("Create-group response: HTTP %s\n%s",
                           resp.status_code, resp.text)

            if resp.ok:
                created = resp.json()
                uuid = created["uuid"]
                self.log.info(
                    "Created scope group '%s'  uuid=%s  "
                    "(filterType=%s  %d IP(s))",
                    group_name, uuid, filter_type, len(ips)
                )
                return uuid

            err_msg = resp.text
            self.log.warning(
                "filterType '%s' rejected (HTTP %s: %s) — trying next",
                filter_type, resp.status_code, err_msg
            )
            last_error = err_msg

        raise RuntimeError(
            f"Could not create IP scope group after trying all filterTypes "
            f"{self._IP_FILTER_CANDIDATES}. Last error: {last_error}"
        )

    def delete_target(self, target_uuid: str) -> bool:
        """DELETE /targets/{uuid} — remove an existing target.

        If target deletion returns 500 (e.g. discovery actively running or referenced by backend),
        logs a warning and returns False so target can be updated or reused.
        """
        url = f"{self.base_url}/targets/{target_uuid}"
        self.log.debug("DELETE %s", url)
        resp = self.session.delete(url)
        self.log.debug("Delete-target response: HTTP %s", resp.status_code)
        if resp.status_code in (200, 204, 404):
            self.log.info("Target %s deleted successfully (HTTP %s)", target_uuid, resp.status_code)
            return True
        self.log.warning(
            "Could not delete existing target UUID=%s (HTTP %s: %s). Will attempt to update/reuse target.",
            target_uuid, resp.status_code, resp.text
        )
        return False

    def update_target(self, target_uuid: str, target_dto: dict) -> dict | None:
        """PUT /targets or PUT /targets/{uuid} — update an existing target's configuration.

        Returns updated target dict if successful, or None if target does not exist or PUT is unsupported.
        """
        # Try PUT /targets/{uuid} first
        url = f"{self.base_url}/targets/{target_uuid}"
        body = json.dumps(target_dto)
        self.log.debug("PUT %s\nRequest body:\n%s", url, self._sanitise_body(body))
        resp = self.session.put(url, data=body)
        self.log.debug("Update-target response: HTTP %s\n%s", resp.status_code, resp.text)
        if resp.ok:
            self.log.info("Target %s updated successfully", target_uuid)
            return resp.json() if resp.text else {"uuid": target_uuid}

        # Try PUT /targets with uuid in DTO
        dto_with_uuid = dict(target_dto)
        dto_with_uuid["uuid"] = target_uuid
        url_collection = f"{self.base_url}/targets"
        resp_coll = self.session.put(url_collection, data=json.dumps(dto_with_uuid))
        if resp_coll.ok:
            self.log.info("Target %s updated via PUT /targets successfully", target_uuid)
            return resp_coll.json() if resp_coll.text else {"uuid": target_uuid}

        self.log.warning("Could not update target %s (HTTP %s): %s", target_uuid, resp.status_code, resp.text)
        return None

    def find_existing_target(self, target_id: str, target_type: str = "WMI") -> dict | None:
        """Check whether a target with the given targetId display name already exists.

        A 504 on POST /targets can occur *after* the server has already persisted
        the target.  Before retrying the POST, call this to avoid duplicates.
        """
        url = f"{self.base_url}/targets"
        self.log.debug("GET %s  (checking for existing %s target '%s')",
                       url, target_type, target_id)
        resp = self.session.get(url)
        if not resp.ok:
            self.log.debug("Could not list targets: HTTP %s", resp.status_code)
            return None
        for t in resp.json():
            if t.get("type", "").upper() != target_type.upper():
                continue
            for field in t.get("inputFields", []):
                if field.get("name") == "targetId" and field.get("value") == target_id:
                    self.log.info(
                        "Found existing target: displayName='%s'  uuid=%s",
                        t.get("displayName"), t.get("uuid")
                    )
                    return t
        return None

    def add_target(self, target_dto: dict, timeout: int = 180,
                   max_retries: int = 3) -> dict:
        """POST /targets — create a new target.

        Uses a *timeout* of 180 s (the nginx gateway on tz1.demo.turbonomic.com
        times out at ~61 s by default; we send with a longer client timeout and
        retry on 504 with exponential backoff).

        Before each retry, checks GET /targets to see if the server already
        persisted the target despite returning 504, avoiding duplicate creation.

        Ref: https://www.ibm.com/docs/en/tarm/8.21.1?topic=endpoints-targets-endpoint
        """
        url = f"{self.base_url}/targets"
        body = json.dumps(target_dto)
        self.log.debug("POST %s\nRequest body:\n%s", url,
                       self._sanitise_body(body))

        # Extract the targetId value for duplicate-check between retries
        target_id_value = next(
            (f["value"] for f in target_dto.get("inputFields", [])
             if f["name"] == "targetId"),
            ""
        )

        for attempt in range(1, max_retries + 1):
            try:
                resp = self.session.post(url, data=body, timeout=timeout)
            except requests.exceptions.Timeout:
                self.log.warning(
                    "POST /targets timed out after %ds (attempt %d/%d)",
                    timeout, attempt, max_retries
                )
                # Check if the server created the target despite the timeout
                existing = self.find_existing_target(target_id_value)
                if existing:
                    self.log.info(
                        "Target was created by the server before timeout. "
                        "Using existing target UUID=%s", existing.get("uuid")
                    )
                    return existing
                if attempt < max_retries:
                    wait = 2 ** attempt
                    self.log.info("Retrying in %ds…", wait)
                    time.sleep(wait)
                    continue
                raise RuntimeError(
                    f"POST /targets timed out after {max_retries} attempts"
                )

            self.log.debug("Add-target response: HTTP %s\n%s",
                           resp.status_code, resp.text)

            if resp.status_code == 504:
                self.log.warning(
                    "POST /targets returned 504 Gateway Timeout (attempt %d/%d). "
                    "Checking if target was created despite the timeout…",
                    attempt, max_retries
                )
                existing = self.find_existing_target(target_id_value)
                if existing:
                    self.log.info(
                        "Target exists (created before gateway timed out). "
                        "UUID=%s", existing.get("uuid")
                    )
                    return existing
                if attempt < max_retries:
                    wait = 2 ** attempt
                    self.log.info("Target not found yet. Retrying in %ds…", wait)
                    time.sleep(wait)
                    continue
                raise RuntimeError(
                    "POST /targets returned 504 after all retries and target "
                    "was not found in GET /targets"
                )

            self._raise_for_status(resp, "add target")
            created = resp.json()
            self.log.info("Target created successfully.  UUID=%s  status=%s",
                          created.get("uuid"), created.get("status"))
            return created

        raise RuntimeError("add_target: exhausted retries without a result")

    def rediscover_target(self, target_uuid: str) -> dict:
        """POST /targets/{uuid}?rediscover=true — trigger a full rediscovery.

        Ref: https://www.ibm.com/docs/en/tarm/8.21.1?topic=endpoints-targets-endpoint
        """
        url = f"{self.base_url}/targets/{target_uuid}?rediscover=true"
        self.log.debug("POST %s  (triggering rediscovery)", url)
        resp = self.session.post(url)
        self.log.debug("Rediscovery response: HTTP %s\n%s",
                       resp.status_code, resp.text)
        self._raise_for_status(resp, "trigger rediscovery")
        result = resp.json()
        self.log.info("Rediscovery triggered for target UUID=%s", target_uuid)
        return result

    def get_target(self, target_uuid: str, retries: int = 3) -> dict:
        """GET /targets/{uuid} — fetch the current state of a target.

        Retries on connection-level errors (RemoteDisconnected, ConnectionError)
        which occur when the server drops an idle keep-alive connection between
        poll intervals.
        """
        url = f"{self.base_url}/targets/{target_uuid}"
        for attempt in range(1, retries + 1):
            try:
                self.log.debug("GET %s  (polling target status)", url)
                resp = self.session.get(url, timeout=30)
                self._raise_for_status(resp, "get target")
                return resp.json()
            except (requests.exceptions.ConnectionError,
                    requests.exceptions.Timeout) as exc:
                if attempt < retries:
                    self.log.debug(
                        "GET %s connection error (attempt %d/%d): %s — retrying",
                        url, attempt, retries, exc
                    )
                    time.sleep(2)
                else:
                    raise

    def poll_discovery_status(self, target_uuid: str,
                              poll_interval: int = 10,
                              max_wait: int = 300) -> str:
        """Poll until the target reaches a terminal health/validation state.

        The Turbonomic API returns two overlapping status indicators:
          - data["status"]                       legacy string: "Validating", "Validated", "Failed"
          - data["healthSummary"]["healthState"]  current enum: "NORMAL", "CRITICAL", "MAJOR", "MINOR"

        This method checks both so it works regardless of which field the
        instance populates.  Terminal conditions:
          success : status == "Validated"  OR  healthState == "NORMAL"
          failure : status == "Failed"     OR  healthState in ("CRITICAL", "MAJOR")
        """
        self.log.info(
            "Polling discovery status for UUID=%s  (interval=%ds, timeout=%ds)",
            target_uuid, poll_interval, max_wait
        )
        elapsed = 0
        while elapsed < max_wait:
            data = self.get_target(target_uuid)

            legacy_status  = data.get("status", "Unknown")
            health_summary = data.get("healthSummary", {})
            health_state   = health_summary.get("healthState", "")
            last_success   = health_summary.get("timeOfLastSuccessfulDiscovery", "")

            self.log.info(
                "  [%3ds elapsed]  status=%-12s  healthState=%s  lastSuccess=%s",
                elapsed, legacy_status, health_state, last_success
            )

            # Success: "Validated", "Discovered", or healthState "NORMAL"
            if legacy_status in ("Validated", "Discovered") or health_state == "NORMAL":
                self.log.info(
                    "Discovery successful.  "
                    "status=%s  healthState=%s  lastSuccess=%s",
                    legacy_status, health_state, last_success
                )
                return "Validated"

            # Failure
            if legacy_status == "Failed" or health_state in ("CRITICAL", "MAJOR"):
                self.log.error(
                    "Discovery failed.  "
                    "status=%s  healthState=%s",
                    legacy_status, health_state
                )
                self.log.debug("Full target response:\n%s",
                               json.dumps(data, indent=2))
                return "Failed"

            time.sleep(poll_interval)
            elapsed += poll_interval

        self.log.warning("Timed out waiting for discovery after %ds", max_wait)
        return "Timeout"

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _raise_for_status(self, resp: requests.Response, operation: str) -> None:
        if not resp.ok:
            self.log.error("%s failed: HTTP %s\n%s",
                           operation, resp.status_code, resp.text)
            raise RuntimeError(
                f"{operation} failed (HTTP {resp.status_code}): {resp.text}"
            )

    @staticmethod
    def _sanitise_body(body: str) -> str:
        """Redact the WMI target password from log output."""
        try:
            obj = json.loads(body)
            for field in obj.get("inputFields", []):
                if field.get("name") == "password":
                    field["value"] = "***REDACTED***"
            return json.dumps(obj, indent=2)
        except Exception:
            return body


# ---------------------------------------------------------------------------
# Target DTO builder
# ---------------------------------------------------------------------------

def extract_windows_targets(config: dict) -> list[dict]:
    """Extract list of individual Windows target configuration dicts.

    Supports:
      1. List of target dicts under 'targets.windows' or 'windows':
         "targets": { "windows": [ { "display_name": "...", "username": "...", "password": "...", "vm_name": "...", ... }, ... ] }
      2. Single target dict with 'vms' list (expands to list with inherited credentials):
         "targets": { "windows": { "username": "...", "password": "...", "vms": [ { "nameOrAddress": "...", "vm_name": "..." }, ... ] } }
      3. Legacy:
         "wmi_target": { ... } or "wmi_targets": [ ... ]
    """
    win_entry = None
    if "targets" in config and isinstance(config["targets"], dict):
        win_entry = config["targets"].get("windows")
    elif "windows" in config:
        win_entry = config["windows"]
    elif "windows_targets" in config:
        win_entry = config["windows_targets"]
    elif "wmi_targets" in config:
        win_entry = config["wmi_targets"]
    elif "wmi_target" in config:
        win_entry = config["wmi_target"]

    if isinstance(win_entry, list):
        return [dict(t) for t in win_entry if isinstance(t, dict)]

    if isinstance(win_entry, dict):
        # If 'vms' array exists inside the single dict, expand each VM to an individual target
        if "vms" in win_entry and isinstance(win_entry["vms"], list):
            expanded_targets = []
            for vm in win_entry["vms"]:
                vm_target = dict(win_entry)
                vm_target.pop("vms", None)
                if isinstance(vm, dict):
                    vm_target.update(vm)
                elif isinstance(vm, str):
                    vm_target["nameOrAddress"] = vm
                    vm_target["vm_name"] = vm
                if not vm_target.get("display_name"):
                    vm_label = vm_target.get("vm_name") or vm_target.get("nameOrAddress") or "Windows-VM"
                    vm_target["display_name"] = f"WMI-{vm_label}"
                expanded_targets.append(vm_target)
            return expanded_targets
        return [win_entry]

    return []


def build_wmi_target_dto(target_cfg: dict, wmi_spec: dict | None = None,
                         scope_uuid: str = "") -> dict:
    """Construct the TargetApiDTO for a WMI target.

    Field names and mandatory/optional status are driven by the live
    /targets/specs WMI entry.  From the actual spec on tz1.demo.turbonomic.com:

      targetEntities  (mandatory, GROUP_SCOPE) – Turbonomic entity or group UUID.
                      Must be a real UUID, NOT a plain IP address.
      targetId        (mandatory, STRING)       – user-visible display label only.
      username        (mandatory, STRING)
      password        (mandatory, STRING)
      domainName      (optional,  STRING)       – leave blank for WORKGROUP/local.
      useNTLM         (optional,  BOOLEAN)      – default true
      secure          (optional,  BOOLEAN)      – HTTPS, default false
      fullValidation  (optional,  BOOLEAN)      – default false

    *scope_uuid* must be a resolved Turbonomic UUID (VM entity UUID from
    /search, or a group UUID from /groups).  Passing a raw IP causes HTTP 500.

    Category is "Guest OS Processes" as documented at:
    https://www.ibm.com/docs/en/tarm/8.21.1?topic=targets-wmi
    """
    t = target_cfg
    log = logging.getLogger("turbo.wmi.dto")

    # Log the full live spec so every field name is visible in the DEBUG log
    if wmi_spec:
        spec_names = [f["name"] for f in wmi_spec.get("inputFields", [])]
        log.debug("Live WMI spec field names: %s", spec_names)

    # For WORKGROUP machines there is no AD domain.
    # The docs say: "Leave blank for local accounts."
    domain_value = "" if t.get("domain", "").upper() in ("WORKGROUP", "") else t.get("domain", "")

    log.debug("targetEntities UUID: %s", scope_uuid)

    display_label = t.get("display_name") or t.get("vm_name") or t.get("nameOrAddress") or "WMI-Windows-Target"

    input_fields: list[dict] = [
        # Mandatory fields
        {"name": "targetEntities", "value": scope_uuid},
        {"name": "targetId",       "value": display_label},
        {"name": "username",       "value": t["username"]},
        {"name": "password",       "value": t["password"]},
        # Optional fields
        {"name": "domainName",     "value": domain_value},
        {"name": "useNTLM",        "value": str(t.get("use_ntlm", True)).lower()},
        {"name": "secure",         "value": str(t.get("use_https", False)).lower()},
        {"name": "fullValidation", "value": "false"},
    ]

    dto = {
        "category":    t.get("category", "Guest OS Processes"),
        "type":        t.get("type", "WMI"),
        "inputFields": input_fields,
    }

    if display_label:
        dto["displayName"] = display_label

    log.debug("Final DTO inputFields: %s",
              [{"name": f["name"],
                "value": "***" if f["name"] == "password" else f["value"]}
               for f in input_fields])
    return dto


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    default_config = (
        Path(__file__).parent.parent / "config" / "config.json"
        if (Path(__file__).parent.parent / "config" / "config.json").exists()
        else Path(__file__).parent / "config.json"
    )

    parser = argparse.ArgumentParser(
        description="Create separate Turbonomic WMI targets for Windows VMs and trigger discovery"
    )
    parser.add_argument(
        "--config",
        default=str(default_config),
        help="Path to the JSON configuration file (default: config/config.json)"
    )
    args = parser.parse_args()

    # --- Load config --------------------------------------------------------
    config_path = Path(args.config)
    if not config_path.is_absolute() and not config_path.exists():
        # Check relative to parent directory (works on both Windows and Unix)
        alt_path = Path(__file__).resolve().parent.parent / args.config
        if alt_path.exists():
            config_path = alt_path

    config_path = config_path.resolve()

    if not config_path.exists():
        print(f"ERROR: Config file not found: {config_path}", file=sys.stderr)
        return 1

    # Read as text first so we can strip a UTF-8 BOM and any trailing
    # whitespace/garbage that editors on Windows sometimes append, which
    # would cause json.load() to raise "Extra data".
    raw = config_path.read_text(encoding="utf-8-sig").strip()
    try:
        config = json.loads(raw)
    except json.JSONDecodeError as exc:
        print(f"ERROR: Failed to parse config file {config_path}: {exc}", file=sys.stderr)
        return 1

    # --- Setup logging (early, so every step is captured) -------------------
    log_cfg = config.get("logging", {})
    # Anchor the log file next to the config so it is predictable on all
    # platforms regardless of what the working directory happens to be.
    log_file_name = log_cfg.get("log_file", "wmi_target_setup.log")
    log_file_path = (
        Path(log_file_name) if Path(log_file_name).is_absolute()
        else config_path.parent / log_file_name
    )
    log = setup_logging(
        level=log_cfg.get("level", "DEBUG"),
        log_file=str(log_file_path)
    )

    windows_targets = extract_windows_targets(config)

    # Check for Linux targets to log notice that they are ignored
    linux_count = 0
    if "targets" in config and isinstance(config["targets"], dict):
        lin_entry = config["targets"].get("linux")
        if isinstance(lin_entry, list):
            linux_count = len(lin_entry)
        elif isinstance(lin_entry, dict):
            linux_count = len(lin_entry.get("vms", [1]))

    log.info("=" * 60)
    log.info("Turbonomic WMI Target Setup (Individual VM Targets)")
    log.info("Config file      : %s", config_path.resolve())
    log.info("Turbo host       : %s", config["turbonomic"]["host"])
    log.info("Windows targets  : %d VM target(s) configured", len(windows_targets))
    if linux_count:
        log.info("Linux target/VMs : %d VM(s) found in config (IGNORED by WMI script)", linux_count)
    log.info("=" * 60)

    if not windows_targets:
        log.error("No Windows target configuration found under 'targets.windows'.")
        return 1

    turbo = TurbonomicClient(
        base_url=config["turbonomic"]["api_base_url"],
        verify_ssl=False   # Turbonomic appliances use self-signed certs by default
    )

    # --- Step 1: Authenticate -----------------------------------------------
    log.info("STEP 1: Authenticating with Turbonomic API")
    turbo_user = config["turbonomic"].get("admin_username") or config["turbonomic"].get("username")
    turbo_pass = config["turbonomic"].get("admin_password") or config["turbonomic"].get("password")
    turbo.login(
        username=turbo_user,
        password=turbo_pass
    )

    # --- Step 2: Verify WMI probe is available ------------------------------
    log.info("STEP 2: Verifying WMI probe availability via /targets/specs")
    wmi_spec = turbo.find_wmi_spec()
    if wmi_spec is None:
        log.error(
            "WMI probe not found in /targets/specs.  "
            "Ensure the WMI probe is enabled in the Turbonomic CR."
        )
        return 1
    log.info("WMI probe is available. Proceeding to process individual Windows target(s).")

    disc_cfg = config.get("discovery", {})
    should_rediscover = disc_cfg.get("trigger_rediscovery_after_add", True)
    overall_success = True

    # --- Step 3: Resolve Single Shared Scope Group for All Windows VMs -------
    log.info("STEP 3: Resolving single shared scope group for all %d Windows VM(s)", len(windows_targets))
    shared_group_name = "Group-WMI-Windows-Fleet"
    all_resolved_vm_uuids: list[str] = []
    missing_vm_configs: list[dict] = []

    for t_cfg in windows_targets:
        vm_name = t_cfg.get("vm_name") or t_cfg.get("display_name") or ""
        ip = t_cfg.get("nameOrAddress") or ""
        log.info("Searching inventory for VM entity: vm_name='%s', IP='%s'", vm_name, ip)
        v_uuid = turbo.search_entity_by_ip(ip, vm_name=vm_name)
        if v_uuid:
            if v_uuid not in all_resolved_vm_uuids:
                all_resolved_vm_uuids.append(v_uuid)
        else:
            missing_vm_configs.append(t_cfg)

    if all_resolved_vm_uuids:
        log.info("Creating/updating single static scope group '%s' (Type=Static) containing all %d VM(s)",
                 shared_group_name, len(all_resolved_vm_uuids))
        shared_scope_uuid = turbo.create_static_vm_scope_group(all_resolved_vm_uuids, shared_group_name)
        if missing_vm_configs:
            log.warning("%d VM(s) could not be resolved from inventory: %s",
                        len(missing_vm_configs),
                        [c.get("vm_name") or c.get("nameOrAddress") for c in missing_vm_configs])
    else:
        log.warning("No VM entities resolved in inventory. Creating static scope group '%s' (Type=Static)", shared_group_name)
        shared_scope_uuid = turbo.create_static_vm_scope_group([], shared_group_name)
        if not shared_scope_uuid:
            all_ips = [t.get("nameOrAddress") for t in windows_targets if t.get("nameOrAddress")]
            shared_scope_uuid = turbo.create_ip_scope_group(all_ips, shared_group_name)

    log.info("Shared targetEntities Scope Group '%s' UUID: %s", shared_group_name, shared_scope_uuid)

    # --- Step 4: Create/Update Individual WMI Target for Each VM using Shared Group ---
    for idx, target_cfg in enumerate(windows_targets, start=1):
        vm_name = target_cfg.get("vm_name") or target_cfg.get("display_name") or ""
        ip = target_cfg.get("nameOrAddress") or ""
        target_display_name = target_cfg.get("display_name") or f"WMI-{vm_name or ip}"
        username = target_cfg.get("username", "")
        domain = target_cfg.get("domain", "")

        log.info("=" * 60)
        log.info("[%d/%d] Configuring WMI Target: '%s'", idx, len(windows_targets), target_display_name)
        log.info("  VM Name: %s | IP: %s | User: %s | Domain: %s", vm_name, ip, username, domain)
        log.info("  Associated Scope Group: '%s' (uuid=%s)", shared_group_name, shared_scope_uuid)
        log.info("=" * 60)

        # Check existing target
        existing = turbo.find_existing_target(target_display_name)
        target_deleted = False
        if existing:
            log.info("Found existing target UUID=%s — attempting to delete before creating with shared static scope group",
                     existing["uuid"])
            target_deleted = turbo.delete_target(existing["uuid"])

        # Override or use the shared scope UUID
        scope_uuid = target_cfg.get("scope_uuid", "").strip() or shared_scope_uuid

        # Build & Add / Update target
        target_dto = build_wmi_target_dto(target_cfg, wmi_spec, scope_uuid)
        log.debug("Target DTO (password redacted):\n%s",
                  TurbonomicClient._sanitise_body(json.dumps(target_dto)))

        target_uuid = None
        if existing and not target_deleted:
            target_uuid = existing["uuid"]
            log.info("Attempting to update existing WMI target UUID=%s", target_uuid)
            updated_target = turbo.update_target(target_uuid, target_dto)
            if updated_target and updated_target.get("uuid"):
                target_uuid = updated_target["uuid"]
                log.info("WMI target updated successfully. UUID=%s", target_uuid)
            else:
                log.info("Update failed. Re-creating via POST /targets")
                created_target = turbo.add_target(target_dto)
                target_uuid = created_target["uuid"]
                log.info("WMI target added successfully. UUID=%s  Initial status=%s",
                         target_uuid, created_target.get("status"))
        else:
            log.info("Adding WMI target '%s' via POST /targets", target_display_name)
            created_target = turbo.add_target(target_dto)
            target_uuid = created_target["uuid"]
            log.info("WMI target added successfully. UUID=%s  Initial status=%s",
                 target_uuid, created_target.get("status"))

        # Trigger rediscovery & Poll
        if should_rediscover and target_uuid:
            log.info("Triggering full rediscovery for target UUID=%s", target_uuid)
            turbo.rediscover_target(target_uuid)

            log.info("Polling discovery status for target UUID=%s", target_uuid)
            final_status = turbo.poll_discovery_status(
                target_uuid=target_uuid,
                poll_interval=disc_cfg.get("rediscover_poll_interval_seconds", 10),
                max_wait=disc_cfg.get("rediscover_max_wait_seconds", 300)
            )
            if final_status == "Validated":
                log.info("SUCCESS: Target '%s' is validated and discovery is complete.", target_display_name)
            elif final_status == "Timeout":
                log.warning("TIMEOUT: Discovery for target '%s' did not complete within wait window.", target_display_name)
                overall_success = False
            else:
                log.error("FAILED: Target '%s' discovery ended with status '%s'.", target_display_name, final_status)
                overall_success = False

    log.info("=" * 60)
    if overall_success:
        log.info("ALL Windows WMI targets processed successfully.")
    else:
        log.warning("Some Windows WMI targets encountered warnings or failures. Check logs for details.")
    log.info("Log written to: %s", log_cfg.get("log_file", "wmi_target_setup.log"))
    log.info("=" * 60)

    return 0 if overall_success else 1


if __name__ == "__main__":
    sys.exit(main())
