import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="koala_xlcalculator-bradbase",
    version="0.0.1b",
    author="Bradley van Ree",
    author_email="brads@bradbase.net",
    description="koala_xlcalcualtor converts MS Excel formulas to Python and evaluates them.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/bradbase/koala_xlcalculator",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        "Operating System :: OS Independent",
    ],
    install_requires=[
            'jsonpickle >= 1.3',
            'networkx >= 2.4',
            'matplotlib >= 3.1.1',
            'numpy >= 1.18.1',
            'pandas >= 1.0.1',
            'openpyxl >= 3.0.3',
        ],
    python_requires='>=3.7.6',
)
