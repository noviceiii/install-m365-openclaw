from setuptools import setup, find_packages

setup(
    name='m365_openclaw',
    version='0.2.0',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'O365',
        'msal',
        'msal_extensions',
        'python-dotenv',
        'python-docx',
        'python-pptx',
        'openpyxl',
        'requests',
    ],
)
