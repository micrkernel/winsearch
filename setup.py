#!/usr/bin/env python
# -*- coding: utf-8 -*-

from setuptools import setup, find_packages
from winsearch import __version__, __author__


setup(
	name = "winsearch",
	version = __version__,
	packages = ["adodbapi"],
	scripts = ["winsearch.py"],

	author = __author__,
	author_email = "jan-magel@web.de",
	description = "",
	license = "MIT",
	url = "https://github.com/micrkernel/winsearch"
	)
