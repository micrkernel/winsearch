winsearch
======

A python library for accessing the windows search database and executing search queries.
This allows a much faster search of files on the system.


Requirements
============

- [adodbapi](https://pypi.python.org/pypi/adodbapi): use `easy_install adodbapi` to install
- Python 2.4 or higher
- Windows Vista or higher


What is Windows Search?
=======================
> Windows Search, formerly known as Windows Desktop Search (WDS) on Windows XP and Windows Server 2003, is an indexed desktop search platform created by Microsoft for Microsoft Windows.
> Windows Search collectively refers to the indexed search on Windows Vista and later versions of Windows (also referred to as Instant Search as well as Windows Desktop Search, a standalone add-on for Windows 2000, Windows XP and Windows Server 2003 made available as freeware. All incarnations of Windows Search share a common architecture and indexing technology and use a compatible application programming interface (API).
For futher information see [Windows Search](http://en.wikipedia.org/wiki/Windows_Search)


Examples
========

Find the five most recently edited documents
--------------------------------------------

	import winsearch
	query = winsearch.Query(select=["path"])
	query.limit(5)
	query.extensions("pdf", "doc", "docx", "odt")
	query.sortby("date").desc()

	cursor = query.execute()
	print "Last five docs:"
	for row in cursor:
		print "\t", row[0]

Get total memory usage of all indexed jpegs
-------------------------------------------

	import winsearch
	query = winsearch.Query(select=["size"])
	query.extensions("jpg", "jpeg")

	cursor = query.execute()
	results = cursor.fetchall()

	total_size = sum(map(lambda r: r[0], results))
	print "Size of all indexed jpegs: %d MBytes" % (total_size // 0x100000)


License
========
The MIT License (MIT)

Copyright (c) 2013 micrkernel

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

