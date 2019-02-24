import os
import sys 
from setuptools import setup

with open('requirements.txt') as f:
    required = f.read().splitlines()

setup(
    name="inetsix-excel-to-template",
    version="0.9",
    scripts=["bin/inetsix-excel-to-template"],
    python_requires=">=2.7",
    install_requires=required,
    url="https://github.com/titom73/python-excel-serializer",
    license="MIT",
    author="Thomas Grimonet",
    author_email="tom@inetsix.net",
    description="Jinja2 template rendering from Excel file",
)
