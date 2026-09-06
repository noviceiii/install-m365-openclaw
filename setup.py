from setuptools import setup, find_packages

setup(
    name='m365_openclaw',
    version='0.6.2',
    description='Microsoft 365 Skill for OpenClaw agents (MSAL hybrid: device-code + application Mail.Send)',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'msal>=1.32,<2',
        'python-dotenv>=1.2.2,<2',
        'python-docx>=1.1,<2',
        'python-pptx>=1.0,<2',
        'requests>=2.34,<3',
    ],
    python_requires='>=3.10',
)
