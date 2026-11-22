import argparse
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timezone
from html import unescape
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
from urllib.error import URLError
from urllib.request import urlopen
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName


INPUT = ("subscription ID", "Region", "Service Name", "SKU", "Required cores")
RESULT = ("Quota increase needed (y/n)", "current quota", "increase to")
PRODUCT_CATALOG = "https://azure.microsoft.com/en-us/products/"
REGION_CATALOG = "https://learn.microsoft.com/en-us/azure/reliability/regions-list"
VM_SERVICES = ("Virtual Machines", "Virtual Machine Scale Sets")
SERVICES = {
    "azure vm", "azure virtual machine", "azure virtual machines", "virtual machine",
    "virtual machines", "vm", "vmss", "virtual machine scale sets", "microsoft.compute",
    "microsoft.compute/virtualmachines", "microsoft.compute/virtualmachinescalesets",
}
GUID = re.compile(r"[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}", re.I)
REGION = re.compile(r"[a-z][a-z0-9]*")
MAX_CORES = 2_147_483_647


class CliError(Exception):
    pass


@dataclass
class Request:
    line: int
    values: list
    family: str = ""
    can_check: bool = False
    issue: str = ""


@dataclass(frozen=True)
class Quota:
    limit: int
    used: int
    source: str


def positive_cores(value):
    if isinstance(value, str):
        value = value.strip()
        value = int(value) if re.fullmatch(r"[0-9]{1,10}", value) else None
    if type(value) is float and value.is_integer():
        value = int(value)
    return value if type(value) is int and 1 <= value <= MAX_CORES else None


def read_requests(path):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book["Requests"] if "Requests" in book else book.worksheets[0]
        rows = sheet.values
        headings = [str(cell).strip().casefold() for cell in next(rows, ())]
        columns = []
        for name in INPUT:
            key = name.casefold()
            if headings.count(key) != 1:
                raise ValueError(f"Header row must contain exactly one '{name}' column.")
            columns.append(headings.index(key))
        requests = []
        for line, row in enumerate(rows, 2):
            values = [row[i].strip() if isinstance(row[i], str) else row[i] for i in columns]
            if all(value is None or value == "" for value in values):
                continue
            request = Request(line, values)
            sub, region, service, sku, cores = values
            count = positive_cores(cores)
            if not isinstance(sub, str) or not GUID.fullmatch(sub):
                request.issue = "Enter a valid Azure subscription ID."
            elif not isinstance(region, str) or not REGION.fullmatch(region.lower()):
                request.issue = "Use an Azure region code, such as eastus or westeurope."
            elif not isinstance(service, str) or service.casefold() not in SERVICES:
                request.issue = "Unsupported service. This report checks standard Azure VM/VMSS vCPUs only."
            elif not isinstance(sku, str) or not sku:
                request.issue = "Enter an Azure VM size, such as Standard_D4s_v5."
            elif count is None:
                request.issue = "Required cores must be a positive whole number."
            else:
                request.can_check = True
                values[4] = count
            requests.append(request)
        if not requests:
            raise ValueError("No requests found. Add at least one row under the headers.")
        return requests
    finally:
        book.close()


def az_json(cli, *args):
    try:
        result = subprocess.run(
            [cli, *args, "--only-show-errors", "--output", "json"],
            capture_output=True, encoding="utf-8", errors="replace", timeout=120, check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise CliError(f"az {' '.join(args[:3])}: {exc}") from exc
    if result.returncode:
        message = " ".join(result.stderr.split()) or f"exit code {result.returncode}"
        raise CliError(f"az {' '.join(args[:3])}: {message[:800]}")
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise CliError(f"az {' '.join(args[:3])} returned invalid JSON.") from exc


def az_list(cli, *args):
    items = az_json(cli, *args)
    if not isinstance(items, list) or not all(isinstance(item, dict) for item in items):
        raise CliError(f"az {' '.join(args[:3])} returned an unexpected list format.")
    return items


def named_numbers(items, get_name, get_value):
    result = {}
    for item in items:
        name, number = get_name(item), get_value(item)
        if isinstance(name, str) and type(number) is int and number >= 0:
            result[name.casefold()] = number
    return result


def field(item, *keys):
    for key in keys:
        if not isinstance(item, dict):
            return None
        item = item.get(key)
    return item


def quota_readings(cli, sub, region, wanted):
    scope = f"/subscriptions/{sub}/providers/Microsoft.Compute/locations/{region}"
    options = ("--scope", scope, "--subscription", sub)
    readings = {}
    try:
        limits = named_numbers(
            az_list(cli, "quota", "list", *options),
            lambda item: item.get("name"),
            lambda item: (
                field(item, "properties", "limit", "value")
                if field(item, "properties", "isQuotaApplicable") is not False else None
            ),
        )
        usage = named_numbers(
            az_list(cli, "quota", "usage", "list", *options),
            lambda item: item.get("name"),
            lambda item: (
                field(item, "properties", "usages", "value")
                if field(item, "properties", "isQuotaApplicable") is not False else None
            ),
        )
        readings.update({name: Quota(limits[name], usage[name], "az quota") for name in limits.keys() & usage.keys() & wanted})
        primary_issue = "az quota did not return both a numeric limit and usage"
    except CliError as exc:
        primary_issue = str(exc)
    missing = wanted - readings.keys()
    if not missing:
        return readings, ""

    try:
        items = az_list(cli, "vm", "list-usage", "--location", region, "--subscription", sub)
        def usage_name(item):
            return field(item, "name", "value")

        limits = named_numbers(items, usage_name, lambda item: item.get("limit"))
        usage = named_numbers(items, usage_name, lambda item: item.get("currentValue"))
        readings.update({
            name: Quota(limits[name], usage[name], "az vm list-usage")
            for name in limits.keys() & usage.keys() & missing
        })
        fallback_issue = "az vm list-usage did not return both a numeric limit and usage"
    except CliError as exc:
        fallback_issue = str(exc)
    return readings, f"{primary_issue}; {fallback_issue}."


def resolve_quotas(cli, requests):
    scopes = defaultdict(list)
    for request in requests:
        if request.can_check:
            sub, region = request.values[:2]
            scopes[sub.lower(), region.lower()].append(request)

    details = {}
    for (sub, region), group in sorted(scopes.items()):
        try:
            skus = az_list(cli, "vm", "list-skus", "--location", region, "--resource-type",
                           "virtualMachines", "--all", "true", "--subscription", sub)
            families = {
                item["name"].casefold(): item["family"].casefold()
                for item in skus
                if isinstance(item.get("name"), str) and isinstance(item.get("family"), str) and item["family"]
            }
            sku_issue = "SKU not found for this region or Azure returned no VM family."
        except CliError as exc:
            families = {}
            sku_issue = str(exc)

        for request in group:
            request.family = families.get(request.values[3].casefold(), "")
            if not request.family:
                request.issue = sku_issue
        wanted = {"cores"} | {request.family for request in group if request.family}
        readings, quota_issue = quota_readings(cli, sub, region, wanted)
        for name in sorted(wanted):
            details[sub, region, name] = readings.get(name)
        for request in group:
            names = ["cores"] + ([request.family] if request.family else [])
            missing = [name for name in names if name not in readings]
            if missing:
                issue = f"Could not verify {', '.join(missing)}: {quota_issue}"
                request.issue = " ".join(filter(None, (request.issue, issue)))
    return details


def table(sheet, headings, widths, rows):
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A2"
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(i)].width = width
    for y, values in enumerate([headings, *rows], 1):
        for x, value in enumerate(values, 1):
            cell = sheet.cell(y, x, value)
            if isinstance(value, str):
                cell.data_type = "s"
            cell.font = Font(name="Arial", size=10, bold=y == 1, color="FFFFFF" if y == 1 else "172B41")
            cell.fill = PatternFill("solid", fgColor="153A5A" if y == 1 else ("EDF5FB" if y % 2 == 0 else "FFFFFF"))
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        sheet.row_dimensions[y].height = 29 if y == 1 else 24
    sheet.auto_filter.ref = sheet.dimensions


def add_guide(book, rows):
    guide = book.create_sheet("Guide")
    table(guide, ("Topic", "How to use this workbook"), (24, 104), rows)
    for row in range(2, guide.max_row + 1):
        guide.row_dimensions[row].height = 34
    guide.auto_filter.ref = None


def download_catalog(url):
    try:
        with urlopen(url, timeout=20) as page:
            return page.read().decode("utf-8")
    except (URLError, TimeoutError) as exc:
        raise ValueError(f"Could not download the Azure catalog at {url}: {exc}") from exc


def label(html):
    return " ".join(unescape(re.sub(r"<[^>]+>", "", html)).split())


def get_service_catalog():
    html = download_catalog(PRODUCT_CATALOG)
    headings = re.findall(r'<h3\b[^>]*\bclass="[^"]*\bh5\b[^"]*"[^>]*>(.*?)</h3>', html, re.S)
    names = {label(heading) for heading in headings}
    names.discard("")
    if len(names) < 100 or not names.issuperset(VM_SERVICES):
        raise ValueError(f"Could not read the Azure product list at {PRODUCT_CATALOG}; the page may have changed.")
    return sorted(names, key=str.casefold)


def get_region_catalog():
    regions = {}
    for table_html in re.findall(r"<table\b[^>]*>(.*?)</table>", download_catalog(REGION_CATALOG), re.S):
        headers = [label(th) for th in re.findall(r"<th\b[^>]*>(.*?)</th>", table_html, re.S)]
        if "Region" not in headers or "Programmatic name" not in headers:
            continue
        name_col, code_col = headers.index("Region"), headers.index("Programmatic name")
        for row in re.findall(r"<tr\b[^>]*>(.*?)</tr>", table_html, re.S):
            cells = re.findall(r"<td\b[^>]*>(.*?)</td>", row, re.S)
            if len(cells) != len(headers):
                continue
            code, name = label(cells[code_col]), label(cells[name_col])
            if name and REGION.fullmatch(code):
                regions[code] = name
    if len(regions) < 40 or not regions.keys() >= {"eastus", "westeurope"}:
        raise ValueError(f"Could not read the Azure region list at {REGION_CATALOG}; the page may have changed.")
    return sorted(regions.items())


def add_dropdown(sheet, column, source, range_name, field_name):
    reference = f"'{source.title}'!$A$2:$A${source.max_row}"
    sheet.parent.defined_names.add(DefinedName(range_name, attr_text=reference))
    validation = DataValidation(
        type="list", formula1=range_name, allow_blank=True, showDropDown=False,
        errorStyle="stop", showErrorMessage=True, errorTitle=f"Unknown Azure {field_name}",
        error=f"Select a {field_name} from the dropdown in this cell.",
    )
    sheet.add_data_validation(validation)
    validation.add(f"{column}2:{column}1048576")


def write_input(path, services, regions, example=False):
    book = Workbook()
    sheet = book.active
    sheet.title = "Requests"
    placeholder = "00000000-0000-0000-0000-000000000000"
    rows = [
        (placeholder, "eastus", VM_SERVICES[0], "Standard_D4s_v5", 16),
        (placeholder, "eastus", VM_SERVICES[0], "Standard_D8s_v5", 16),
        (placeholder, "westeurope", VM_SERVICES[1], "Standard_E4s_v5", 20),
    ] if example else []
    table(sheet, INPUT, (40, 20, 54, 25, 18), rows)
    catalog = book.create_sheet("Azure services")
    table(catalog, ("Service Name", "vCPU quota report"), (56, 26),
          [(name, "Supported" if name in VM_SERVICES else "Not checked") for name in services])
    add_dropdown(sheet, "C", catalog, "AzureServiceNames", "service")
    region_list = book.create_sheet("Azure regions")
    table(region_list, ("Region code", "Region name"), (26, 32), regions)
    add_dropdown(sheet, "B", region_list, "AzureRegionCodes", "region")
    validation = DataValidation(type="whole", operator="between", formula1=1, formula2=MAX_CORES,
                                allow_blank=True, showErrorMessage=True, errorTitle="Invalid vCPU count",
                                error="Enter a positive whole number.")
    sheet.add_data_validation(validation)
    validation.add("E2:E1048576")
    add_guide(book, [
        ("Start here", "Enter your requests on the Requests tab. Replace all example subscription IDs with actual IDs."),
        ("Region", "Select a region code. The Azure regions tab maps codes to names. Region access can vary by subscription."),
        ("Region list source", REGION_CATALOG),
        ("SKU", "Use an exact VM size (Standard_D4s_v5)."),
        ("Service Name", "Select from the dropdown. Only Virtual Machines and Virtual Machine Scale Sets are checked for standard vCPU quotas."),
        ("Catalog source", f"Public Azure products when this template was generated: {PRODUCT_CATALOG}"),
        ("Required cores", "Enter the total additional vCPUs for this row. Example: four 4-vCPU VMs = 16."),
        ("Shared quotas", "Rows in the same subscription and region are added together to assess the whole deployment plan."),
    ])
    book.save(path)


def write_report(path, requests, details):
    book = Workbook()
    report = book.active
    report.title = "Quota report"
    detail = book.create_sheet("Quota details")
    issues = book.create_sheet("Issues")
    index = {key: i for i, key in enumerate(details, 2)}
    scope_keys = [
        (request.values[0].lower(), request.values[1].lower()) if request.can_check else ()
        for request in requests
    ]
    links = [(index.get((*scope, req.family)), index.get((*scope, "cores")))
             for req, scope in zip(requests, scope_keys)]
    table(report, INPUT + RESULT,
          (40, 18, 22, 24, 16, 27, 16, 52),
          [(*req.values, None, None, None) for req in requests])
    if requests:
        plan = book.create_sheet("Plan data")
        plan.append(("Family bucket", "Region bucket", "Required cores"))
        for req, (family, regional) in zip(requests, links):
            plan.append((family, regional, req.values[4] if req.can_check else None))
        plan.sheet_state = "hidden"
    report["G1"].comment = Comment("Current quota is the configured VM-family vCPU limit.", "Azure quota report")
    report["H1"].comment = Comment("Minimum new quota limit(s), not the number of cores to add.", "Azure quota report")
    table(detail, ("subscription ID", "Region", "Quota resource", "Current limit", "Current usage",
                   "Plan vCPUs", "Increase to", "Azure CLI source"),
          (40, 18, 30, 17, 17, 15, 16, 24),
          [(*key, quota.limit if quota else None, quota.used if quota else None, None, None,
            quota.source if quota else "Not verified") for key, quota in details.items()])
    detail.auto_filter.ref = None
    last = len(requests) + 1
    plan_ref = "'Plan data'!"
    for row, key in enumerate(details, 2):
        bucket = "B" if key[2] == "cores" else "A"
        detail[f"F{row}"] = f"=SUMIF({plan_ref}${bucket}$2:${bucket}${last},{row},{plan_ref}$C$2:$C${last})"
        detail[f"G{row}"] = f'=IF(AND(COUNT(D{row}:E{row})=2,E{row}+F{row}>D{row}),E{row}+F{row},"")'

    for row, (family, regional) in enumerate(links, 2):
        if regional is None:
            continue
        ref = "'Quota details'!"
        f_target = f"{ref}$G${family}" if family is not None else '""'
        r_target = f"{ref}$G${regional}"
        f_ok = f"COUNT({ref}$D${family}:$E${family})=2" if family is not None else "FALSE"
        r_ok = f"COUNT({ref}$D${regional}:$E${regional})=2"
        f_need, r_need = f"ISNUMBER({f_target})", f"ISNUMBER({r_target})"
        report[f"F{row}"] = f'=IF(OR({f_need},{r_need}),"y",IF(AND({f_ok},{r_ok}),"n",""))'
        if family is not None:
            limit = f"{ref}$D${family}"
            report[f"G{row}"] = f'=IF(ISNUMBER({limit}),{limit},"")'
        report[f"H{row}"] = (f'=IF({f_need},"VM family: "&{f_target},"")'
                            f'&IF(AND({f_need},{r_need}),"; ","")'
                            f'&IF({r_need},"Regional: "&{r_target},"")')

    table(issues, ("Input row", "Report row", "Issue"), (15, 15, 112),
          [(req.line, row, req.issue) for row, req in enumerate(requests, 2) if req.issue])
    for row in range(2, issues.max_row + 1):
        issues.row_dimensions[row].height = 64
    add_guide(book, [
        ("Snapshot (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC") if requests else "Blank output template"),
        ("Quota increase needed", "y = at least one limit must increase; n = both checked limits cover the plan. Blank = could not determine."),
        ("Current quota", "The VM-family limit. See Quota details for current usage and the total regional limit (resource: cores)."),
        ("Increase to", "Minimum new limit = current usage + planned additional vCPUs for that quota. No increase is requested automatically."),
        ("Shared quotas", "The plan sums rows sharing a VM family or total regional limit. Repeated targets are not separate increases."),
        ("Issues and capacity", "Fix any Issues and rerun: missing inputs can undercount the plan. Sufficient quota does not guarantee SKU availability or capacity."),
    ])
    book.save(path)


def main():
    parser = argparse.ArgumentParser(description="Report standard Azure VM/VMSS vCPU quotas from an Excel deployment plan.")
    parser.add_argument("input", nargs="?", type=Path, help="Input .xlsx workbook")
    parser.add_argument("-o", "--output", type=Path, default=Path("quota_report.xlsx"), help="Output .xlsx workbook")
    parser.add_argument("--make-templates", action="store_true", help="Write blank input/output templates and an example input")
    args = parser.parse_args()

    try:
        if args.make_templates:
            if args.input is not None:
                parser.error("Do not pass an input workbook with --make-templates.")
            services = get_service_catalog()
            regions = get_region_catalog()
            write_input(Path("quota_input_template.xlsx"), services, regions)
            write_input(Path("quota_input_example.xlsx"), services, regions, example=True)
            write_report(Path("quota_output_template.xlsx"), [], {})
            return 0
        if args.input is None:
            parser.error("Provide an input .xlsx workbook or use --make-templates.")
        if args.input.suffix.lower() != ".xlsx" or args.output.suffix.lower() != ".xlsx":
            raise ValueError("Input and output must be .xlsx files.")
        if args.input.resolve() == args.output.resolve() or (args.output.exists() and args.input.samefile(args.output)):
            raise ValueError("The output cannot overwrite the input workbook.")
        requests = read_requests(args.input)
        cli = shutil.which("az")
        if cli is None:
            raise CliError("Azure CLI not found on PATH. Install Azure CLI and run az login first.")
        try:
            az_json(cli, "account", "show")
        except CliError as exc:
            raise CliError(f"Could not use Azure CLI login. Run az login first. {exc}") from exc
        details = resolve_quotas(cli, requests)
        write_report(args.output, requests, details)
        errors = sum(bool(request.issue) for request in requests)
        print(f"Wrote {args.output} ({len(requests)} request(s)).")
        if errors:
            print(f"{errors} row(s) could not be fully verified. See the Issues tab.", file=sys.stderr)
        return 1 if errors else 0
    except (CliError, ValueError, OSError, InvalidFileException, BadZipFile, ParseError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
