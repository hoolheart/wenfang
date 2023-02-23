from setuptools import setup, find_packages
import codecs

# prepare description
description = 'Utilities for automation in general office work'
long_description = description
with codecs.open('README.rst', 'r', encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='wenfang',
    version='0.0.1',
    license='MIT',
    description=description,
    long_description=long_description,
    author='Edward',
    requires=[],
    packages=find_packages(include=['wenfang', 'wenfang.*']),
    classifiers=[
        'Intended Audience :: Developers',
        'Programming Language :: Python :: 3',
        'Topic :: Software Developemnt :: Libraries',
        'Topic :: Office Automation',
    ],
)
