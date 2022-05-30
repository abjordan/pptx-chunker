from setuptools import setup, find_packages

with open('README.md') as f:
    readme = f.read()

with open('LICENSE') as f:
    license = f.read()

setup(
    name='pptxchunker',
    version='0.1.0',
    description='Split up PPTX files',
    long_description=readme,
    author='Alex Jordan',
    author_email='abjordan@gmail.com',
    url='https://github.com/abjordan/pptx-chunker',
    license=license,
    packages=find_packages(exclude=('tests', 'docs'))
)