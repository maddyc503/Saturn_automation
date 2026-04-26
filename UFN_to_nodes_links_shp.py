from __future__ import annotations

import subprocess
import shutil
import logging
import csv
from pathlib import Path

log = logging.getLogger("p1x_runner")


def setup_logger(output_dir: Path) -> Path:
    """
    Configure logging to write to both the console and a timestamped log file
    in output_dir. Returns the path to the log file.
    """
    import datetime
    log.setLevel(logging.DEBUG)

    # Console handler — INFO and above, clean format (no timestamp)
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("%(message)s"))

    # File handler — DEBUG and above, full format with timestamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = output_dir / f"p1x_run_{timestamp}.log"
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s",
                          datefmt="%Y-%m-%d %H:%M:%S")
    )

    log.addHandler(console)
    log.addHandler(file_handler)

    log.info("Log file: %s", log_path)
    return log_path


P1X_VERSIONS = {
    "1": (r"C:\Program Files (x86)\Atkins\SATURN\XEXES 11.4.07H MC N4\$P1X.exe", "11.4.07H MC N4"),
    "2": (r"C:\Program Files (x86)\Atkins\SATURN\XEXES 11.5.05N MC N4\$P1X.exe", "11.5.05N MC N4"),
}

CLEANUP_PATTERNS = ["*.VDU", "*.LPX", "*.CTL", "*.LOG"]

# Link attributes CSV output filename — must match what is hardcoded in the SATDB key
LINK_CSV_NAME = "Link_attributes.csv"

# Headers for the link attributes CSV (columns in order of appearance)
LINK_HEADERS = [
    "Node_X",
    "Node_Y",
    "Lanes",
    "Capacity_Index",
    "Free_flow_speed",
    "Distance",
]

# Common coordinate systems grouped by region: (display label, EPSG code)
CRS_OPTIONS = [
    # --- British ---
    ("British National Grid (OSGB 1936)           - EPSG:27700", 27700),
    ("OSGB 1936 / British National Grid + ODN ht  - EPSG:7405",  7405),
    # --- Irish ---
    ("Irish Transverse Mercator (ITM)              - EPSG:2157",  2157),
    ("TM65 / Irish Grid (older datum)              - EPSG:29902", 29902),
    ("TM75 / Irish Grid                            - EPSG:29903", 29903),
    # --- Indian ---
    ("WGS 84 geographic (lat-lon, global/India)    - EPSG:4326",  4326),
    ("WGS 84 / UTM zone 43N  (NW India)            - EPSG:32643", 32643),
    ("WGS 84 / UTM zone 44N  (N-central India)     - EPSG:32644", 32644),
    ("WGS 84 / UTM zone 45N  (NE India)            - EPSG:32645", 32645),
    ("Kalianpur 1975 / India zone I   (NW)         - EPSG:24378", 24378),
    ("Kalianpur 1975 / India zone IIa (NW-central) - EPSG:24379", 24379),
    ("Kalianpur 1975 / India zone IIb (NE-central) - EPSG:24380", 24380),
    ("Kalianpur 1975 / India zone III (S-central)  - EPSG:24381", 24381),
    ("Kalianpur 1975 / India zone IV  (S)          - EPSG:24382", 24382),
    # --- Other UTM ---
    ("WGS 84 / UTM zone 29N                        - EPSG:32629", 32629),
    ("WGS 84 / UTM zone 30N                        - EPSG:32630", 32630),
    ("WGS 84 / UTM zone 31N                        - EPSG:32631", 32631),
    # --- Manual ---
    ("Enter EPSG code manually", None),
]

# ---------------------------------------------------------------------------
# Key file templates
# ---------------------------------------------------------------------------

NODES_KEY_TEMPLATE = """\
      1721        79   70    0  Files menu     (Pixels/key/status/Line)     6001
      1772        91   79    0  Outputs        (Pixels/key/status/Line)     6105
      1762       110   88    0  X,Y co-ords    (Pixels/key/status/Line)     6105
{output_dir}
Yes Currently XYUNIT =      10.0; do you want to output new X,Y co-ordinates
 Menu bar:  &Exit                             1700
Yes OK to quit the program?
`
\
"""

SATDB_KEY_TEMPLATE = """\
      1765       419   66    0  SATDB Opts     (Pixels/key/status/Line)     6001
           2                                                                6800
           6                                                                7052
           2                                                                7070
           3                                                                7070
           5                                                                7070
           0                                                                7070
           0                                                                7095
           4                                                                6800
           1                                                                6820
           1                                                                6801
           3                                                                6801
           5                                                                6801
           6                                                                6801
           0                                                                6801
          13                                                                6800
           0                                                                7530
{link_csv_path}
           0                                                                6800
      1492       235         0 19073.8 9285.53 (Mouse pixels/status/X,Y)    6001
 Menu bar:  &Exit                             1700
Yes OK to quit the program?
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def prompt_choice(prompt: str, options: list[str]) -> int:
    """Print numbered options and return a validated 0-based index."""
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    while True:
        raw = input(prompt).strip()
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return int(raw) - 1
        print(f"  Please enter a number between 1 and {len(options)}.")


def prompt_yes_no(prompt: str) -> bool:
    """Ask a yes/no question and return True for yes."""
    while True:
        raw = input(prompt).strip().lower()
        if raw in ("yes", "y"):
            return True
        if raw in ("no", "n"):
            return False
        print("  Please enter yes or no.")


def glob_unique(directory: Path, *patterns: str) -> list[Path]:
    """Glob multiple patterns, deduplicating by case-insensitive name."""
    seen: set[str] = set()
    results: list[Path] = []
    for pattern in patterns:
        for f in sorted(directory.glob(pattern)):
            key = f.name.upper()
            if key not in seen:
                seen.add(key)
                results.append(f)
    return results


# ---------------------------------------------------------------------------
# Setup prompts
# ---------------------------------------------------------------------------

def select_version() -> tuple[str, str]:
    log.info("Select SATURN version:")
    labels = [label for _, label in P1X_VERSIONS.values()]
    idx = prompt_choice("Select version (number): ", labels)
    exe, label = P1X_VERSIONS[str(idx + 1)]
    return exe, label


def select_model_dir() -> Path:
    while True:
        raw = input("\nEnter model folder path: ").strip().strip('"')
        path = Path(raw)
        if path.is_dir():
            return path
        log.warning("  Folder not found — please try again.")


def select_ufn_file(directory: Path) -> Path:
    files = glob_unique(directory, "*.UFN", "*.ufn")
    if not files:
        raise FileNotFoundError(f"No UFN files found in {directory}")
    log.info("\nUFN files found:")
    idx = prompt_choice("Select UFN (number): ", [f.name for f in files])
    return files[idx]


def resolve_output_dir(model_dir: Path) -> Path:
    answer = input(
        "\nSave outputs in a new sub-folder? (yes/no): ").strip().lower()
    if answer == "yes":
        raw = input("Enter folder name [default: Outputs]: ").strip()
        name = raw if raw else "Outputs"
        out = model_dir / name
        out.mkdir(exist_ok=True)
        log.info("  Output folder: %s", out)
        return out
    return model_dir


def select_operation() -> tuple[str, bool, bool]:
    """
    Ask which operation to run.
    Returns (operation, create_node_shp, create_link_shp).
    """
    log.info("\nSelect operation to run:")
    options = [
        "Node CSV & Node Shapefile            — extract nodes, save CSV + point shapefile",
        "Link CSV & Link Shapefile            — extract links, save CSV + polyline shapefile",
        "Node CSV, Link CSV & Both Shapefiles — extract everything, save all outputs",
    ]
    idx = prompt_choice("Select operation (number): ", options)

    if idx == 0:
        return "nodes", True, False
    elif idx == 1:
        return "links_only_shp", False, True
    else:
        return "both", True, True


# ---------------------------------------------------------------------------
# Key file writers
# ---------------------------------------------------------------------------

def create_nodes_key_file(temp_dir: Path, ufn_stem: str, output_dir: Path) -> Path:
    """Write the node XY KEY file."""
    nodes_dir = output_dir / "Nodes"
    nodes_dir.mkdir(exist_ok=True)
    key_content = NODES_KEY_TEMPLATE.format(output_dir=str(nodes_dir))
    key_path = temp_dir / f"{ufn_stem}.KEY"
    key_path.write_text(key_content, encoding="utf-8")
    shutil.copy(key_path, output_dir / key_path.name)
    log.info(f"  Created nodes KEY file: {key_path.name}")
    log.info(f"  Output path in KEY: {nodes_dir}")
    nodes_dir.rmdir()
    return key_path


def create_satdb_key_file(temp_dir: Path, ufn_stem: str, output_dir: Path) -> Path:
    """Write the SATDB link attributes KEY file."""
    link_csv_path = output_dir / LINK_CSV_NAME
    key_content = SATDB_KEY_TEMPLATE.format(link_csv_path=str(link_csv_path))
    key_path = temp_dir / f"{ufn_stem}_SATDB.KEY"
    key_path.write_text(key_content, encoding="utf-8")
    shutil.copy(key_path, output_dir / key_path.name)
    log.info(f"  Created SATDB KEY file: {key_path.name}")
    log.info(f"  Link attributes output: {link_csv_path}")
    return key_path


# ---------------------------------------------------------------------------
# Node XY -> CSV conversion
# ---------------------------------------------------------------------------

def convert_xy_to_csv(output_dir: Path) -> Path | None:
    """
    Find the .XY file, filter rows, split by whitespace, and write a headed CSV.

    Headers: Nodes, X_Coordinates, Y_Coordinates

    Filtering rules:
      - Skip the first row (header/title line).
      - Skip the last non-blank row.
      - Skip blank lines.
      - Skip any row whose first non-whitespace character is a letter (a-z / A-Z).
    """
    xy_files = list(output_dir.glob("*.XY")) + list(output_dir.glob("*.xy"))
    if not xy_files:
        log.warning(
            "  No .XY file found in %s — skipping CSV conversion.", output_dir)
        return None

    xy_path = xy_files[0]
    csv_path = output_dir / (xy_path.stem + ".csv")

    all_lines = xy_path.read_text(
        encoding="utf-8", errors="replace").splitlines()

    last_nonblank_idx = None
    for i in range(len(all_lines) - 1, -1, -1):
        if all_lines[i].strip():
            last_nonblank_idx = i
            break

    kept_rows = 0
    skipped_rows = 0

    with csv_path.open("w", newline="", encoding="utf-8") as out_fh:
        writer = csv.writer(out_fh)
        writer.writerow(["Nodes", "X_Coordinates", "Y_Coordinates"])

        for line_num, line in enumerate(all_lines):
            stripped = line.strip()

            if line_num == 0:
                skipped_rows += 1
                continue
            if not stripped:
                skipped_rows += 1
                continue
            if line_num == last_nonblank_idx:
                skipped_rows += 1
                continue
            if stripped[0].isalpha():
                skipped_rows += 1
                continue

            fields = stripped.split()
            if len(fields) < 3:
                log.warning("  Skipping short row (line %d): %s",
                            line_num + 1, stripped)
                skipped_rows += 1
                continue

            writer.writerow(fields[:3])
            kept_rows += 1

    log.info(
        "  Converted %s -> %s  (%d data rows written, %d rows skipped)",
        xy_path.name, csv_path.name, kept_rows, skipped_rows,
    )
    return csv_path


# ---------------------------------------------------------------------------
# Link attributes CSV processing
# ---------------------------------------------------------------------------

def process_link_attributes_csv(output_dir: Path) -> Path | None:
    """
    Read the raw Link_attributes.csv produced by SATDB, split each row by
    whitespace, assign the 6 column headers, and overwrite the file in place.

    Headers: Node_X, Node_Y, Lanes, Capacity_Index, Free_flow_speed, Distance

    Filtering rules (same as nodes):
      - Skip the first row.
      - Skip the last non-blank row.
      - Skip blank lines.
      - Skip any row whose first non-whitespace character is a letter.
    """
    link_csv = output_dir / LINK_CSV_NAME
    if not link_csv.exists():
        log.warning(
            "  %s not found — skipping link attribute processing.", LINK_CSV_NAME)
        return None

    all_lines = link_csv.read_text(
        encoding="utf-8", errors="replace").splitlines()

    last_nonblank_idx = None
    for i in range(len(all_lines) - 1, -1, -1):
        if all_lines[i].strip():
            last_nonblank_idx = i
            break

    kept_rows = 0
    skipped_rows = 0
    n_cols = len(LINK_HEADERS)

    # Write to a temp file alongside then replace
    tmp_path = link_csv.with_suffix(".tmp")
    with tmp_path.open("w", newline="", encoding="utf-8") as out_fh:
        writer = csv.writer(out_fh)
        writer.writerow(LINK_HEADERS)

        for line_num, line in enumerate(all_lines):
            stripped = line.strip()

            if line_num == 0:
                skipped_rows += 1
                continue
            if not stripped:
                skipped_rows += 1
                continue
            if line_num == last_nonblank_idx:
                skipped_rows += 1
                continue
            if stripped[0].isalpha():
                skipped_rows += 1
                continue

            fields = stripped.split()
            if len(fields) < n_cols:
                log.warning(
                    "  Skipping short row (line %d, only %d fields): %s",
                    line_num + 1, len(fields), stripped,
                )
                skipped_rows += 1
                continue

            writer.writerow(fields[:n_cols])
            kept_rows += 1

    # Replace original with processed version
    tmp_path.replace(link_csv)

    log.info(
        "  Processed %s  (%d data rows written, %d rows skipped)",
        LINK_CSV_NAME, kept_rows, skipped_rows,
    )
    return link_csv


# ---------------------------------------------------------------------------
# Shapefile creation
# ---------------------------------------------------------------------------

def select_crs() -> int:
    """Prompt user to pick a coordinate system; returns an EPSG integer."""
    log.info("\nSelect coordinate system for the shapefile:")
    labels = [label for label, _ in CRS_OPTIONS]
    idx = prompt_choice("Select CRS (number): ", labels)
    _, epsg = CRS_OPTIONS[idx]

    if epsg is None:
        while True:
            raw = input("  Enter EPSG code (e.g. 27700): ").strip()
            if raw.isdigit() and int(raw) > 0:
                return int(raw)
            print("  Please enter a valid positive integer EPSG code.")
    return epsg


def select_shapefile_name(output_dir: Path, default: str = "Output") -> Path:
    """Prompt user for a shapefile name (with a default). Returns the full stem path."""
    while True:
        raw = input(
            f"\nEnter shapefile name (without extension) [default: {default}]: ").strip()
        name = raw if raw else default
        if name:
            stem = Path(name).stem
            return output_dir / stem
        print("  Name cannot be empty — please try again.")


def _write_prj(prj_path: Path, epsg: int) -> None:
    """Write a .prj sidecar file with WKT for the chosen EPSG."""
    try:
        from pyproj import CRS
        wkt = CRS.from_epsg(epsg).to_wkt()
        prj_path.write_text(wkt, encoding="utf-8")
        log.info("  Written .prj via pyproj for EPSG:%d", epsg)
        return
    except Exception:
        pass

    FALLBACK_WKT: dict[int, str] = {
        27700: (
            'PROJCS["British_National_Grid",'
            'GEOGCS["GCS_OSGB_1936",'
            'DATUM["D_OSGB_1936",'
            'SPHEROID["Airy_1830",6377563.396,299.3249646]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",400000],'
            'PARAMETER["False_Northing",-100000],'
            'PARAMETER["Central_Meridian",-2],'
            'PARAMETER["Scale_Factor",0.9996012717],'
            'PARAMETER["Latitude_Of_Origin",49],'
            'UNIT["Meter",1]]'
        ),
        2157: (
            'PROJCS["Irish_Transverse_Mercator",'
            'GEOGCS["GCS_GRS_1980",'
            'DATUM["D_GRS_1980",'
            'SPHEROID["GRS_1980",6378137,298.257222101]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",600000],'
            'PARAMETER["False_Northing",750000],'
            'PARAMETER["Central_Meridian",-8],'
            'PARAMETER["Scale_Factor",0.99982],'
            'PARAMETER["Latitude_Of_Origin",53.5],'
            'UNIT["Meter",1]]'
        ),
        29902: (
            'PROJCS["TM65_Irish_Grid",'
            'GEOGCS["GCS_TM65",'
            'DATUM["D_TM65",'
            'SPHEROID["Airy_Modified",6377340.189,299.3249646]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",200000],'
            'PARAMETER["False_Northing",250000],'
            'PARAMETER["Central_Meridian",-8],'
            'PARAMETER["Scale_Factor",1.000035],'
            'PARAMETER["Latitude_Of_Origin",53.5],'
            'UNIT["Meter",1]]'
        ),
        29903: (
            'PROJCS["TM75_Irish_Grid",'
            'GEOGCS["GCS_TM75",'
            'DATUM["D_TM75",'
            'SPHEROID["Airy_Modified",6377340.189,299.3249646]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",200000],'
            'PARAMETER["False_Northing",250000],'
            'PARAMETER["Central_Meridian",-8],'
            'PARAMETER["Scale_Factor",1.000035],'
            'PARAMETER["Latitude_Of_Origin",53.5],'
            'UNIT["Meter",1]]'
        ),
        4326: (
            'GEOGCS["GCS_WGS_1984",'
            'DATUM["D_WGS_1984",'
            'SPHEROID["WGS_1984",6378137,298.257223563]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]]'
        ),
        32629: (
            'PROJCS["WGS_1984_UTM_Zone_29N",'
            'GEOGCS["GCS_WGS_1984",'
            'DATUM["D_WGS_1984",'
            'SPHEROID["WGS_1984",6378137,298.257223563]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",500000],'
            'PARAMETER["False_Northing",0],'
            'PARAMETER["Central_Meridian",-9],'
            'PARAMETER["Scale_Factor",0.9996],'
            'PARAMETER["Latitude_Of_Origin",0],'
            'UNIT["Meter",1]]'
        ),
        32630: (
            'PROJCS["WGS_1984_UTM_Zone_30N",'
            'GEOGCS["GCS_WGS_1984",'
            'DATUM["D_WGS_1984",'
            'SPHEROID["WGS_1984",6378137,298.257223563]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",500000],'
            'PARAMETER["False_Northing",0],'
            'PARAMETER["Central_Meridian",-3],'
            'PARAMETER["Scale_Factor",0.9996],'
            'PARAMETER["Latitude_Of_Origin",0],'
            'UNIT["Meter",1]]'
        ),
        32631: (
            'PROJCS["WGS_1984_UTM_Zone_31N",'
            'GEOGCS["GCS_WGS_1984",'
            'DATUM["D_WGS_1984",'
            'SPHEROID["WGS_1984",6378137,298.257223563]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",500000],'
            'PARAMETER["False_Northing",0],'
            'PARAMETER["Central_Meridian",3],'
            'PARAMETER["Scale_Factor",0.9996],'
            'PARAMETER["Latitude_Of_Origin",0],'
            'UNIT["Meter",1]]'
        ),
        32643: (
            'PROJCS["WGS_1984_UTM_Zone_43N",'
            'GEOGCS["GCS_WGS_1984",'
            'DATUM["D_WGS_1984",'
            'SPHEROID["WGS_1984",6378137,298.257223563]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",500000],'
            'PARAMETER["False_Northing",0],'
            'PARAMETER["Central_Meridian",75],'
            'PARAMETER["Scale_Factor",0.9996],'
            'PARAMETER["Latitude_Of_Origin",0],'
            'UNIT["Meter",1]]'
        ),
        32644: (
            'PROJCS["WGS_1984_UTM_Zone_44N",'
            'GEOGCS["GCS_WGS_1984",'
            'DATUM["D_WGS_1984",'
            'SPHEROID["WGS_1984",6378137,298.257223563]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",500000],'
            'PARAMETER["False_Northing",0],'
            'PARAMETER["Central_Meridian",81],'
            'PARAMETER["Scale_Factor",0.9996],'
            'PARAMETER["Latitude_Of_Origin",0],'
            'UNIT["Meter",1]]'
        ),
        32645: (
            'PROJCS["WGS_1984_UTM_Zone_45N",'
            'GEOGCS["GCS_WGS_1984",'
            'DATUM["D_WGS_1984",'
            'SPHEROID["WGS_1984",6378137,298.257223563]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Transverse_Mercator"],'
            'PARAMETER["False_Easting",500000],'
            'PARAMETER["False_Northing",0],'
            'PARAMETER["Central_Meridian",87],'
            'PARAMETER["Scale_Factor",0.9996],'
            'PARAMETER["Latitude_Of_Origin",0],'
            'UNIT["Meter",1]]'
        ),
        24378: (
            'PROJCS["Kalianpur_1975_India_zone_I",'
            'GEOGCS["GCS_Kalianpur_1975",'
            'DATUM["D_Kalianpur_1975",'
            'SPHEROID["Everest_1830_1975_Definition",6377299.151,300.8017255]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Lambert_Conformal_Conic"],'
            'PARAMETER["False_Easting",2743196.4],'
            'PARAMETER["False_Northing",914398.8],'
            'PARAMETER["Central_Meridian",68],'
            'PARAMETER["Standard_Parallel_1",32.5],'
            'PARAMETER["Scale_Factor",0.99878641],'
            'PARAMETER["Latitude_Of_Origin",39.5],'
            'UNIT["Meter",1]]'
        ),
        24379: (
            'PROJCS["Kalianpur_1975_India_zone_IIa",'
            'GEOGCS["GCS_Kalianpur_1975",'
            'DATUM["D_Kalianpur_1975",'
            'SPHEROID["Everest_1830_1975_Definition",6377299.151,300.8017255]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Lambert_Conformal_Conic"],'
            'PARAMETER["False_Easting",2743196.4],'
            'PARAMETER["False_Northing",914398.8],'
            'PARAMETER["Central_Meridian",74],'
            'PARAMETER["Standard_Parallel_1",26],'
            'PARAMETER["Scale_Factor",0.99878641],'
            'PARAMETER["Latitude_Of_Origin",32.5],'
            'UNIT["Meter",1]]'
        ),
        24380: (
            'PROJCS["Kalianpur_1975_India_zone_IIb",'
            'GEOGCS["GCS_Kalianpur_1975",'
            'DATUM["D_Kalianpur_1975",'
            'SPHEROID["Everest_1830_1975_Definition",6377299.151,300.8017255]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Lambert_Conformal_Conic"],'
            'PARAMETER["False_Easting",2743196.4],'
            'PARAMETER["False_Northing",914398.8],'
            'PARAMETER["Central_Meridian",90],'
            'PARAMETER["Standard_Parallel_1",26],'
            'PARAMETER["Scale_Factor",0.99878641],'
            'PARAMETER["Latitude_Of_Origin",32.5],'
            'UNIT["Meter",1]]'
        ),
        24381: (
            'PROJCS["Kalianpur_1975_India_zone_III",'
            'GEOGCS["GCS_Kalianpur_1975",'
            'DATUM["D_Kalianpur_1975",'
            'SPHEROID["Everest_1830_1975_Definition",6377299.151,300.8017255]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Lambert_Conformal_Conic"],'
            'PARAMETER["False_Easting",2743196.4],'
            'PARAMETER["False_Northing",914398.8],'
            'PARAMETER["Central_Meridian",80],'
            'PARAMETER["Standard_Parallel_1",19],'
            'PARAMETER["Scale_Factor",0.99878641],'
            'PARAMETER["Latitude_Of_Origin",26],'
            'UNIT["Meter",1]]'
        ),
        24382: (
            'PROJCS["Kalianpur_1975_India_zone_IV",'
            'GEOGCS["GCS_Kalianpur_1975",'
            'DATUM["D_Kalianpur_1975",'
            'SPHEROID["Everest_1830_1975_Definition",6377299.151,300.8017255]],'
            'PRIMEM["Greenwich",0],'
            'UNIT["Degree",0.017453292519943295]],'
            'PROJECTION["Lambert_Conformal_Conic"],'
            'PARAMETER["False_Easting",2743196.4],'
            'PARAMETER["False_Northing",914398.8],'
            'PARAMETER["Central_Meridian",80],'
            'PARAMETER["Standard_Parallel_1",12],'
            'PARAMETER["Scale_Factor",0.99878641],'
            'PARAMETER["Latitude_Of_Origin",19],'
            'UNIT["Meter",1]]'
        ),
    }

    wkt = FALLBACK_WKT.get(epsg)
    if wkt:
        prj_path.write_text(wkt, encoding="utf-8")
        log.info("  Written .prj using built-in WKT for EPSG:%d", epsg)
    else:
        log.warning(
            "  pyproj is not installed and no built-in WKT exists for EPSG:%d.\n"
            "  Install pyproj (pip install pyproj) for full CRS support.\n"
            "  The .prj file was NOT written — set the CRS manually in your GIS.", epsg
        )


def create_shapefile(csv_path: Path, shp_stem: Path, epsg: int) -> bool:
    """
    Read the nodes CSV and write a Point shapefile using pyshp.
    Returns True on success.
    """
    try:
        import shapefile
    except ImportError:
        log.error(
            "  pyshp is not installed. Run:  pip install pyshp\n"
            "  Then re-run the script to create the shapefile."
        )
        return False

    rows: list[tuple[int, float, float]] = []
    with csv_path.open("r", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            try:
                node = int(row["Nodes"])
                x = float(row["X_Coordinates"])
                y = float(row["Y_Coordinates"])
                rows.append((node, x, y))
            except (ValueError, KeyError) as exc:
                log.warning("  Skipping malformed CSV row: %s (%s)", row, exc)

    if not rows:
        log.error("  No valid rows found in CSV — shapefile not created.")
        return False

    with shapefile.Writer(str(shp_stem), shapeType=shapefile.POINT) as w:
        w.field("Nodes",   "N", size=10)
        w.field("X_Coord", "F", decimal=3)
        w.field("Y_Coord", "F", decimal=3)
        for node, x, y in rows:
            w.point(x, y)
            w.record(node, x, y)

    _write_prj(shp_stem.with_suffix(".prj"), epsg)
    log.info(
        "  Shapefile created: %s.shp  (%d points, EPSG:%d)",
        shp_stem.name, len(rows), epsg,
    )
    return True


def load_node_lookup(nodes_csv: Path) -> dict[int, tuple[float, float]]:
    """
    Read the nodes CSV and return a dict mapping node_id -> (X, Y).
    Used to look up coordinates when building the links shapefile.
    """
    lookup: dict[int, tuple[float, float]] = {}
    with nodes_csv.open("r", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            try:
                node_id = int(row["Nodes"])
                x = float(row["X_Coordinates"])
                y = float(row["Y_Coordinates"])
                lookup[node_id] = (x, y)
            except (ValueError, KeyError) as exc:
                log.warning("  Skipping malformed node row: %s (%s)", row, exc)
    log.info("  Loaded %d nodes from %s", len(lookup), nodes_csv.name)
    return lookup


def create_links_shapefile(
    link_csv: Path,
    nodes_csv: Path,
    shp_stem: Path,
    epsg: int,
) -> bool:
    """
    Build a Polyline shapefile from the link attributes CSV.

    Each row contains Node_X and Node_Y (node IDs, not coordinates).
    Coordinates are looked up from the nodes CSV and used as the
    start and end points of each link LineString.

    All 6 link attributes are written as shapefile fields.
    Links where either node is missing from the lookup are skipped with a warning.
    """
    try:
        import shapefile
    except ImportError:
        log.error(
            "  pyshp is not installed. Run:  pip install pyshp\n"
            "  Then re-run the script to create the shapefile."
        )
        return False

    # Build node coordinate lookup from the nodes CSV
    node_lookup = load_node_lookup(nodes_csv)
    if not node_lookup:
        log.error("  Node lookup is empty — cannot build links shapefile.")
        return False

    written = 0
    skipped = 0

    with shapefile.Writer(str(shp_stem), shapeType=shapefile.POLYLINE) as w:
        # Define attribute fields matching LINK_HEADERS
        w.field("Node_X",     "N", size=10)          # start node ID
        w.field("Node_Y",     "N", size=10)          # end node ID
        w.field("Lanes",      "N", size=5)
        # Capacity_Index (truncated to 10 chars)
        w.field("Cap_Idx",    "N", size=10)
        w.field("FF_Speed",   "F", decimal=3)        # Free_flow_speed
        w.field("Distance",   "F", decimal=3)

        with link_csv.open("r", encoding="utf-8") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                try:
                    node_x_id = int(row["Node_X"])
                    node_y_id = int(row["Node_Y"])
                    lanes = int(row["Lanes"])
                    cap_idx = int(row["Capacity_Index"])
                    ff_speed = float(row["Free_flow_speed"])
                    distance = float(row["Distance"])
                except (ValueError, KeyError) as exc:
                    log.warning(
                        "  Skipping malformed link row: %s (%s)", row, exc)
                    skipped += 1
                    continue

                # Look up start and end coordinates from nodes
                if node_x_id not in node_lookup:
                    log.warning(
                        "  Node %d (Node_X) not found in nodes CSV — skipping link %d->%d",
                        node_x_id, node_x_id, node_y_id,
                    )
                    skipped += 1
                    continue
                if node_y_id not in node_lookup:
                    log.warning(
                        "  Node %d (Node_Y) not found in nodes CSV — skipping link %d->%d",
                        node_y_id, node_x_id, node_y_id,
                    )
                    skipped += 1
                    continue

                x1, y1 = node_lookup[node_x_id]
                x2, y2 = node_lookup[node_y_id]

                # Write polyline: a single part with two vertices
                w.line([[[x1, y1], [x2, y2]]])
                w.record(node_x_id, node_y_id, lanes,
                         cap_idx, ff_speed, distance)
                written += 1

    _write_prj(shp_stem.with_suffix(".prj"), epsg)
    log.info(
        "  Links shapefile created: %s.shp  (%d links written, %d skipped, EPSG:%d)",
        shp_stem.name, written, skipped, epsg,
    )
    return True


# ---------------------------------------------------------------------------
# Link polyline shapefile
# ---------------------------------------------------------------------------

def load_node_coords(nodes_csv: Path) -> dict[int, tuple[float, float]]:
    """
    Read the nodes CSV and return a dict mapping node_id -> (x, y).
    Expects columns: Nodes, X_Coordinates, Y_Coordinates.
    """
    coords: dict[int, tuple[float, float]] = {}
    with nodes_csv.open("r", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            try:
                node_id = int(row["Nodes"])
                x = float(row["X_Coordinates"])
                y = float(row["Y_Coordinates"])
                coords[node_id] = (x, y)
            except (ValueError, KeyError) as exc:
                log.warning("  Skipping malformed node row: %s (%s)", row, exc)
    log.info("  Loaded coordinates for %d nodes.", len(coords))
    return coords


def create_link_shapefile(
    link_csv: Path,
    nodes_csv: Path,
    shp_stem: Path,
    epsg: int,
) -> bool:
    """
    Build a Polyline shapefile from Link_attributes.csv using node coordinates
    looked up from the nodes CSV.

    Each link row:  Node_X (A-node), Node_Y (B-node), Lanes, Capacity_Index,
                    Free_flow_speed, Distance
    Geometry: straight line from A-node coords -> B-node coords.
    All 6 attributes are stored on each feature.
    """
    try:
        import shapefile
    except ImportError:
        log.error(
            "  pyshp is not installed. Run:  pip install pyshp\n"
            "  Then re-run the script to create the link shapefile."
        )
        return False

    # Load node lookup table
    node_coords = load_node_coords(nodes_csv)

    # Read link rows
    links = []
    missing_nodes: set[int] = set()

    with link_csv.open("r", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            try:
                a_node = int(row["Node_X"])
                b_node = int(row["Node_Y"])
                lanes = float(row["Lanes"])
                cap_index = float(row["Capacity_Index"])
                ffs = float(row["Free_flow_speed"])
                distance = float(row["Distance"])
            except (ValueError, KeyError) as exc:
                log.warning("  Skipping malformed link row: %s (%s)", row, exc)
                continue

            # Look up coordinates for both ends
            if a_node not in node_coords:
                missing_nodes.add(a_node)
                continue
            if b_node not in node_coords:
                missing_nodes.add(b_node)
                continue

            ax, ay = node_coords[a_node]
            bx, by = node_coords[b_node]
            links.append((a_node, b_node, ax, ay, bx, by,
                          lanes, cap_index, ffs, distance))

    if missing_nodes:
        log.warning(
            "  %d node(s) referenced in links but not found in nodes CSV: %s%s",
            len(missing_nodes),
            ", ".join(str(n) for n in sorted(missing_nodes)[:10]),
            " ..." if len(missing_nodes) > 10 else "",
        )

    log.info("  Link rows read: %d valid, %d missing nodes skipped.",
             len(links), len(missing_nodes))
    if not links:
        log.error("  No valid links to write — link shapefile not created.")
        log.error(
            "  Check that Node_X/Node_Y values in the link CSV match node numbers in the nodes CSV.")
        return False

    with shapefile.Writer(str(shp_stem), shapeType=shapefile.POLYLINE) as w:
        # Attribute fields
        w.field("Node_X",         "N", size=10)
        w.field("Node_Y",         "N", size=10)
        w.field("Lanes",          "F", decimal=2)
        w.field("Cap_Index",      "F", decimal=2)
        w.field("FFS",            "F", decimal=2)
        w.field("Distance",       "F", decimal=4)

        for (a_node, b_node, ax, ay, bx, by,
             lanes, cap_index, ffs, distance) in links:
            # A straight two-point polyline
            w.line([[[ax, ay], [bx, by]]])
            w.record(a_node, b_node, lanes, cap_index, ffs, distance)

    _write_prj(shp_stem.with_suffix(".prj"), epsg)

    log.info(
        "  Link shapefile created: %s.shp  (%d links, EPSG:%d)",
        shp_stem.name, len(links), epsg,
    )
    return True


# ---------------------------------------------------------------------------
# Cleanup
# ---------------------------------------------------------------------------

def auto_cleanup(temp_files: list[Path]) -> None:
    """
    Silently delete all temporary/intermediate files (.KEY, .XY, etc.).
    These are never needed after the run completes.
    """
    if not temp_files:
        return
    log.info("\n--- Auto-cleanup (temp files) ---")
    deleted = 0
    for p in temp_files:
        if p and p.exists():
            try:
                p.unlink()
                log.info("  Deleted: %s", p.name)
                deleted += 1
            except Exception as exc:
                log.error("  Could not delete %s: %s", p.name, exc)
    log.info("  %d temp file(s) removed.", deleted)


def _run_p1x_process(exe: str, ufn_file: Path, key_file: Path,
                     temp_dir: Path, output_dir: Path) -> None:
    """Run the P1X executable and move outputs to output_dir."""
    input_names = {ufn_file.name.upper(), key_file.name.upper()}

    result = subprocess.run(
        [exe, ufn_file.stem, key_file.name, "vdu", "v"],
        cwd=temp_dir,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        log.error("P1X stderr:\n%s", result.stderr)
        raise RuntimeError(f"P1X exited with code {result.returncode}")

    for pattern in CLEANUP_PATTERNS:
        for f in temp_dir.glob(pattern):
            f.unlink()

    for f in temp_dir.iterdir():
        if f.name.upper() not in input_names:
            dest = output_dir / f.name
            if dest.exists():
                log.warning("  Overwriting existing file: %s", dest.name)
            shutil.move(str(f), dest)


def run_nodes(
    exe: str,
    ufn_file: Path,
    output_dir: Path,
    model_dir: Path,
    create_shp: bool = True,
) -> tuple[Path | None, Path | None]:
    """Run node XY extraction. Returns (xy_path, csv_path)."""
    temp_dir = model_dir / "TEMP_RUN_NODES"
    temp_dir.mkdir(exist_ok=True)
    try:
        shutil.copy(ufn_file, temp_dir / ufn_file.name)
        key_file = create_nodes_key_file(temp_dir, ufn_file.stem, output_dir)

        log.info(f"\nRunning P1X node extraction ({ufn_file.stem})...")
        _run_p1x_process(exe, ufn_file, key_file, temp_dir, output_dir)
        log.info("  Node extraction complete.")

        log.info("\nConverting .XY file to CSV...")
        csv_path = convert_xy_to_csv(output_dir)

        xy_files = list(output_dir.glob("*.XY")) + \
            list(output_dir.glob("*.xy"))
        xy_path = xy_files[0] if xy_files else None

        if csv_path and create_shp:
            log.info("\nCreating node point shapefile...")
            shp_stem = select_shapefile_name(output_dir, default="Nodes")
            epsg = select_crs()
            create_shapefile(csv_path, shp_stem, epsg)
        elif not create_shp:
            log.info("  Skipping node shapefile creation.")

        return xy_path, csv_path

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def run_links(
    exe: str,
    ufn_file: Path,
    output_dir: Path,
    model_dir: Path,
    nodes_csv: Path | None = None,
    create_shp: bool = True,
) -> Path | None:
    """
    Run SATDB link attribute extraction.
    If nodes_csv is provided and create_shp is True, also creates a polyline shapefile.
    Returns link_csv_path.
    """
    temp_dir = model_dir / "TEMP_RUN_LINKS"
    temp_dir.mkdir(exist_ok=True)
    try:
        shutil.copy(ufn_file, temp_dir / ufn_file.name)
        key_file = create_satdb_key_file(temp_dir, ufn_file.stem, output_dir)

        log.info(f"\nRunning P1X SATDB link extraction ({ufn_file.stem})...")
        _run_p1x_process(exe, ufn_file, key_file, temp_dir, output_dir)
        log.info("  Link extraction complete.")

        log.info("\nProcessing Link_attributes.csv...")
        link_csv = process_link_attributes_csv(output_dir)

        # Create polyline shapefile if requested and we have node coordinates
        if link_csv and create_shp:
            if nodes_csv and nodes_csv.exists():
                log.info("\nCreating link polyline shapefile...")
                log.info("  Using nodes CSV: %s", nodes_csv)
                log.info("  Using link CSV:  %s", link_csv)
                shp_stem = select_shapefile_name(output_dir, default="Links")
                epsg = select_crs()
                success = create_link_shapefile(
                    link_csv, nodes_csv, shp_stem, epsg)
                if not success:
                    log.error(
                        "  Link shapefile creation failed — check warnings above.")
            else:
                log.warning(
                    "  No nodes CSV available — skipping link shapefile.\n"
                    "  Run node extraction first (or choose \'Both\') to enable link shapefiles."
                )
        elif link_csv and not create_shp:
            log.info("  Skipping link shapefile creation.")

        return link_csv

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    exe, label = select_version()
    log.info(f"Using SATURN {label}")

    model_dir = select_model_dir()
    ufn_file = select_ufn_file(model_dir)
    output_dir = resolve_output_dir(model_dir)
    # Start file logger now that we know where outputs go
    log_path = setup_logger(output_dir)

    operation, create_node_shp, create_link_shp = select_operation()

    xy_path: Path | None = None
    nodes_csv: Path | None = None
    link_csv: Path | None = None

    if operation in ("nodes", "both"):
        # Full node extraction — CSV and shapefile saved as user outputs
        xy_path, nodes_csv = run_nodes(
            exe, ufn_file, output_dir, model_dir, create_shp=create_node_shp)

    elif operation == "links_only_shp":
        # Silent node extraction — only used to provide coordinates for link geometry
        log.info("\nExtracting node coordinates for link geometry...")
        xy_path, nodes_csv = run_nodes(
            exe, ufn_file, output_dir, model_dir, create_shp=False)

    if operation in ("links_only_shp", "both"):
        link_csv = run_links(exe, ufn_file, output_dir, model_dir,
                             nodes_csv=nodes_csv, create_shp=create_link_shp)

        # For links_only_shp, remove the intermediate nodes CSV (not a user output)
        if operation == "links_only_shp" and nodes_csv and nodes_csv.exists():
            try:
                nodes_csv.unlink()
                log.debug("  Removed intermediate nodes CSV.")
            except Exception:
                pass
            nodes_csv = None

    # --- Auto-delete temp files (.XY, .KEY) ---
    temp_files: list[Path] = []
    if xy_path:
        temp_files.append(xy_path)
    for pattern in ("*.KEY", "*.dat", "*.DAT", "*.LPP", "*.lpp"):
        for f in output_dir.glob(pattern):
            temp_files.append(f)
    auto_cleanup(temp_files)

    log.info("\nAll done. Outputs in: %s", output_dir)
    log.info("Full log saved to:  %s", log_path)


if __name__ == "__main__":
    main()
