#!/usr/bin/env python3
"""Setup script for nd2ImagesToPowerpoint."""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="nd2images-to-powerpoint",
    version="1.0.0",
    author="miuraTakashi",
    description="Convert Nikon ND2 fluorescence images to PowerPoint presentations",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/miuraTakashi/nd2ImagesToPowerpoint",
    py_modules=["nd2ImagesToPowerpoint"],
    scripts=["nd2ImagesToPowerpoint"],
    install_requires=requirements,
    python_requires=">=3.9",
    entry_points={
        "console_scripts": [
            "nd2ImagesToPowerpoint=nd2ImagesToPowerpoint:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Science/Research",
        "Topic :: Scientific/Engineering :: Bio-Informatics",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
)

