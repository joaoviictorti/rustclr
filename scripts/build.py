import subprocess
import os
import shutil

PROJECT_NAME = "clr"
PROJECT_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "../cli"))
TARGET_DIR = os.path.join(PROJECT_PATH, "target")
OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "bin"))

targets = {
    "x86_64-pc-windows-gnu":  "gnu-x64",
    "i686-pc-windows-gnu":    "gnu-x86",
    "x86_64-pc-windows-msvc": "msvc-x64",
    "i686-pc-windows-msvc":   "msvc-x86"
}

def build_and_copy(target, suffix):
    print(f"[+] Building for {target}...")

    env = os.environ.copy()
    env["RUSTFLAGS"] = (
        "--remap-path-prefix=C:\\Users\\VICTOR=."
        " -C target-feature=+crt-static"
    )

    try:
        subprocess.run(
            ["cargo", "build", "--release", "--target", target],
            cwd=PROJECT_PATH,
            check=True,
            env=env
        )
    except subprocess.CalledProcessError:
        print(f"[!] Failed to build for {target}")
        return

    binary_name = f"{PROJECT_NAME}.exe"
    source = os.path.join(TARGET_DIR, target, "release", binary_name)
    dest_name = f"{PROJECT_NAME}-{suffix}.exe"
    dest_path = os.path.join(OUTPUT_DIR, dest_name)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    try:
        shutil.copyfile(source, dest_path)
        print(f"[✓] Copied to {dest_path}")
    except FileNotFoundError:
        print(f"[!] Binary not found at {source}")

if __name__ == "__main__":
    for target, suffix in targets.items():
        build_and_copy(target, suffix)