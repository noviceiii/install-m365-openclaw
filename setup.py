from setuptools import setup, find_packages

setup(
    name='m365_openclaw',
    version='0.5.0',
    description='Microsoft m365 Skill for OpenClaw agents (delegated / device-code auth)',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'msal',
        'python-dotenv',
        'python-docx',
        'python-pptx',
        'openpyxl',
        'requests',
    ],
    python_requires='>=3.8',
)
