import subprocess
import sys


def run_tests():
    """Run test suite with coverage reporting"""
    cmd = [
        "pytest",
        "--cov=src/",
        "--cov-report=term-missing",
        "--cov-report=html",
        "--benchmark-only" if "--benchmark" in sys.argv else "",
        "-v"
    ]

    result = subprocess.run([arg for arg in cmd if arg])
    return result.returncode


if __name__ == "__main__":
    sys.exit(run_tests())