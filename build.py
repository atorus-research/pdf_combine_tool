import os
import shutil
import PyInstaller.__main__


def clean_build():
    """Clean build directories"""
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
    print("Cleaned build directories")


def build_exe():
    """Build executable using PyInstaller"""
    PyInstaller.__main__.run([
        'main.spec',
        '--clean',
        '--noconfirm'
    ])


def verify_assets():
    """Verify all required assets exist"""
    required_assets = [
        os.path.join('assets', 'images', 'atorus_logo.png'),
        os.path.join('assets', 'images', 'pdf_utility_logo.png'),
        os.path.join('assets', 'images', 'python_logo.png'),
        os.path.join('assets', 'images', 'pdf.ico')
    ]

    missing_assets = []
    for asset in required_assets:
        if not os.path.exists(asset):
            missing_assets.append(asset)

    if missing_assets:
        print("Missing required assets:")
        for asset in missing_assets:
            print(f"  - {asset}")
        raise FileNotFoundError("Missing required assets")

    print("All required assets found")


def main():
    """Main build process"""
    try:
        print("Starting build process...")
        verify_assets()
        clean_build()
        build_exe()
        print("Build completed successfully!")
    except Exception as e:
        print(f"Build failed: {str(e)}")
        exit(1)


if __name__ == '__main__':
    main()