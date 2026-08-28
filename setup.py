from setuptools import setup, find_packages

setup(
    name='m365_openclaw',
    version='0.6.1',
    description='Microsoft m365 Skill for OpenClaw agents (delegated / device-code auth)',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'msal>=1.32,<2',
        'python-dotenv>=1.0.1,<2',
        'python-docx>=1.1,<2',
        'python-pptx>=1.0,<2',
        'requests>=2.32,<3',
    ],
    python_requires='>=3.8',
)
