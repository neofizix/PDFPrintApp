from setuptools import setup, find_packages

setup(
    name="PDFPrinterApp",
    version="1.0.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "Flask",
        "Flask-CORS",
        "pystray",
        "Pillow",
        "pywin32"
    ],
)
