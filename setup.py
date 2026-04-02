from setuptools import setup, find_packages

setup(
    name='m365_openclaw',
    version='0.2.0',
    description='Microsoft m365 Skill for OpenClaw agents',
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
    python_requires='>=3.8',
)
