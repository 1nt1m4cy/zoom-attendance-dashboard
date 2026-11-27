#!/usr/bin/env python3
"""
Attendance Companion — Updater
Run this to pull the latest version from GitHub and reinstall dependencies.
"""
import subprocess, sys, os

REPO_URL = "https://github.com/1nt1m4cy/zoom-attendance-dashboard.git"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def run(cmd, cwd=None):
    print(f"  > {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=cwd or SCRIPT_DIR, capture_output=True, text=True)
    if result.stdout.strip():
        print(result.stdout.strip())
    if result.returncode != 0:
        if result.stderr.strip():
            print(result.stderr.strip())
        return False
    return True

def main():
    print("\n=== Attendance Companion — Update ===\n")

    # Check if we're inside a git repo
    is_git = os.path.isdir(os.path.join(SCRIPT_DIR, ".git"))

    if is_git:
        print("[1/3] Pulling latest changes...")
        # Stash any local changes to .env or other user files
        run(["git", "stash", "--include-untracked", "-m", "auto-stash before update"])
        ok = run(["git", "pull", "--rebase", "origin", "master"])
        if not ok:
            print("\n  Git pull failed. Trying hard reset to origin/master...")
            run(["git", "fetch", "origin"])
            run(["git", "reset", "--hard", "origin/master"])
        # Restore stashed changes (.env etc)
        restore = subprocess.run(
            ["git", "stash", "pop"],
            cwd=SCRIPT_DIR, capture_output=True, text=True
        )
        if restore.returncode != 0 and "No stash entries" not in restore.stderr:
            print("  Note: could not auto-restore stashed files. Run 'git stash pop' manually.")
    else:
        print("[1/3] Not a git repo — cloning fresh...")
        parent = os.path.dirname(SCRIPT_DIR)
        tmp = os.path.join(parent, "_zah_update_tmp")
        run(["git", "clone", REPO_URL, tmp])
        # Copy new files over (skip .env so user creds are preserved)
        for f in os.listdir(tmp):
            if f in (".git",):
                continue
            src = os.path.join(tmp, f)
            dst = os.path.join(SCRIPT_DIR, f)
            if f == ".env" and os.path.exists(dst):
                print(f"  Skipping .env (preserving your credentials)")
                continue
            if os.path.isfile(src):
                import shutil
                shutil.copy2(src, dst)
                print(f"  Updated: {f}")
        # Move .git in so future updates use git pull
        import shutil
        git_src = os.path.join(tmp, ".git")
        git_dst = os.path.join(SCRIPT_DIR, ".git")
        if os.path.isdir(git_src) and not os.path.isdir(git_dst):
            shutil.move(git_src, git_dst)
        shutil.rmtree(tmp, ignore_errors=True)

    print("\n[2/3] Installing/updating dependencies...")
    run([sys.executable, "-m", "pip", "install", "-q", "-r",
         os.path.join(SCRIPT_DIR, "requirements.txt")])

    print("\n[3/3] Ensuring Playwright browser is installed...")
    run([sys.executable, "-m", "playwright", "install", "chromium"])

    print("\n=== Update complete. Run 'python companion.py' to start. ===\n")

if __name__ == "__main__":
    main()
