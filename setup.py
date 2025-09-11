from pathlib import Path
from setuptools import setup

README = (Path(__file__).parent / "README.md").read_text(encoding="utf-8")

setup(
    name="guibatchtranslator",
    version="0.1.0",
    description="Offline batch document translator (Argos Translate + PyQt5), with Excel support",
    long_description=README,
    long_description_content_type="text/markdown",
    author="Justin Hammitt",
    url="https://github.com/JustinHammitt/GUIBatchTranslator",
    license="MIT",
    python_requires=">=3.11,<3.12",  # keep to 3.11 wheels
    # single-module layout
    py_modules=["GUIBatchTranslate"],
    install_requires=[
        "PyQt5",
        "argostranslate==1.9.6",
        "argos-translate-files",
        "sentencepiece==0.2.0",
        "ctranslate2>=4,<5",
        "openpyxl",
        "xlrd==1.2.0",
        "et_xmlfile",
    ],
    # create launchers on install
    entry_points={
        "gui_scripts": [
            "GUIBatchTranslator=GUIBatchTranslate:main",   # no console window on Windows
        ],
        "console_scripts": [
            "guibatchtranslate=GUIBatchTranslate:main",    # useful for debugging from terminal
        ],
    },
    include_package_data=True,  # allows MANIFEST.in to ship extra files if present
    classifiers=[
        "Programming Language :: Python :: 3",
        "Environment :: Win32 (MS Windows)",
        "License :: OSI Approved :: MIT License",
        "Topic :: Text Processing :: Linguistic",
        "Intended Audience :: End Users/Desktop",
    ],
)
