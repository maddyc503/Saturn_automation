import subprocess
import shutil
from pathlib import Path

P1X_VERSIONS = {
    "1": r"C:\Program Files (x86)\Atkins\SATURN\XEXES 11.4.07H MC N4\$P1X.exe",
    "2": r"C:\Program Files (x86)\Atkins\SATURN\XEXES 11.5.05N MC N4\$P1X.exe",
}

# --- SATURN VERSION ---
print("Select SATURN version:")
print("  1. 11.4.07H MC N4")
print("  2. 11.5.05N MC N4")
version_choice = input("Select version (number): ").strip()
if version_choice not in P1X_VERSIONS:
    raise Exception("Invalid version selection")
P1X_EXE = P1X_VERSIONS[version_choice]

# --- FOLDER ---
model_dir = Path(input("\nEnter model folder path: ").strip().strip('"'))
if not model_dir.is_dir():
    raise Exception("Folder not found")

# --- SCAN UFN FILES ---
ufn_files = list(model_dir.glob("*.UFN")) + list(model_dir.glob("*.ufn"))
if not ufn_files:
    raise Exception("No UFN files found in folder")

print("\nUFN files found:")
for i, f in enumerate(ufn_files, 1):
    print(f"  {i}. {f.name}")
ufn_choice = int(input("Select UFN (number): ")) - 1
ufn_file = ufn_files[ufn_choice]

# --- SCAN KEY FILES ---
key_files = list(model_dir.glob("*.KEY")) + list(model_dir.glob("*.key"))
if not key_files:
    raise Exception("No KEY files found in folder")

print("\nKEY files found:")
for i, f in enumerate(key_files, 1):
    print(f"  {i}. {f.name}")
key_choice = int(input("Select KEY (number): ")) - 1
key_file = key_files[key_choice]

# --- OUTPUT FOLDER ---
save_new = input("\nSave outputs in a new folder? (yes/no): ").strip().lower()
if save_new == "yes":
    folder_name = input("Enter folder name: ").strip()
    output_dir = model_dir / folder_name
    output_dir.mkdir(exist_ok=True)
else:
    output_dir = model_dir

# --- RUN ---
temp_dir = model_dir / "TEMP_RUN"
temp_dir.mkdir(exist_ok=True)

input_files = {ufn_file.name.upper(), key_file.name.upper()}

try:
    shutil.copy(ufn_file, temp_dir / ufn_file.name)
    shutil.copy(key_file, temp_dir / key_file.name)

    print(f"\nRunning P1X with {ufn_file.stem} and {key_file.name}...")
    subprocess.run(
        [P1X_EXE, ufn_file.stem, key_file.name, "vdu", "v"],
        cwd=temp_dir
    )

    for pattern in ["*.VDU", "*.LPX", "*.CTL", "*.LOG"]:
        for f in temp_dir.glob(pattern):
            f.unlink()

    for f in temp_dir.iterdir():
        if f.name.upper() not in input_files:
            shutil.move(str(f), output_dir / f.name)

    print(f"Done. Outputs saved in: {output_dir}")

finally:
    shutil.rmtree(temp_dir, ignore_errors=True)
