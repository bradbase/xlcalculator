import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="xlcalculator",
    version="0.0.7b",
    author="Bradley van Ree",
    author_email="brads@bradbase.net",
    description="xlcalcualtor converts MS Excel formulas to Python and evaluates them.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/bradbase/xlcalculator",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Operating System :: OS Independent",
    ],
    install_requires=[
            'jsonpickle >= 1.3',
            'networkx >= 2.4',
            'numpy >= 1.18.1',
            'pandas >= 1.0.1',
            'openpyxl >= 3.0.3',
            'numpy_financial >= 1.0.0',
            'xlfunctions >= 0.0.3b'
        ],
    python_requires='>=3.7.6',
)
