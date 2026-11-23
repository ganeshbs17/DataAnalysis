# PyInstaller spec: one-folder, windowed build of the Notice Generator.
# Build from the repository root:  pyinstaller packaging/NoticeGenerator.spec --noconfirm
from PyInstaller.utils.hooks import collect_data_files

ROOT = SPECPATH + "/.."

datas = [
    (ROOT + "/notice_generator/resources/bank_names.json", "notice_generator/resources"),
    (ROOT + "/notice_generator/resources/officer_profiles.json", "notice_generator/resources"),
]
datas += collect_data_files("docx")      # python-docx default template parts
datas += collect_data_files("docxtpl")

excludes = [
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineQuick",
    "PySide6.Qt3DCore", "PySide6.QtQuick", "PySide6.QtQml", "PySide6.QtMultimedia",
    "PySide6.QtCharts", "PySide6.QtDataVisualization", "PySide6.QtBluetooth",
    "pandas", "numpy", "matplotlib", "tkinter", "pytest", "cryptography",
]

a = Analysis(
    [ROOT + "/packaging/launcher.py"],
    pathex=[ROOT],
    datas=datas,
    hiddenimports=["PySide6.QtPdf", "PySide6.QtPdfWidgets"],
    excludes=excludes,
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="NoticeGenerator",
    console=False,
    upx=False,
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="NoticeGenerator")
