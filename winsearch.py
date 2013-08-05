#!/usr/bin/env python
# -*- coding: utf-8 -*-
import adodbapi as OleDb
import urllib2


__author__ = "microkernel"
__version__ = "0.0.1dev"


CONN_STRING = 'Provider=Search.CollatorDSO;' \
              'Extended Properties=\"Application=Windows\"'
SYSTEMINDEX_COLS = {
    "author": "System.ItemAuthors",
    "date": "System.ItemDate",
    # The extension of a file, without the dot at
    # the beginning
    "fileext": "System.ItemType",
    "path": "System.ItemPathDisplay",
    # Windows Search intern rank of the file
    "rank": "System.Search.Rank",
    # Mimetype of a file
    "mimetype": "MimeType",
    "url": "System.ItemUrl",
    # Filename: Same like os.path.basename(path)
    "name": "System.ItemName",
    # Size of a file in Bytes
    "size": "System.Size",
    # unknowen function
    "kind": "System.Kind",
    # ...more...
}


class Query():
    def __init__(self, select=[], **kwargs):
        self.select(*select or ["System.ItemUrl"])
        self.limit(kwargs.get("limit", 0))  # if 0: no limit
        self.scope(kwargs.get("scope", ""))
        self.extensions(*kwargs.get("extensions", []))
        self.fpattern(kwargs.get("filepattern", ""))
        self.sortby(kwargs.get("sort", None))

        sort_order = kwargs.get("order", "asc").lower()
        if sort_order in ["asc", "desc"]:
            getattr(self, sort_order)()
        else:
            raise ValueError, "Unknowen sort order. Please use desc or asc."

    def __str__(self):
        return self.query

    def __repr__(self):
        return self.query

    @property
    def query(self):
        return self.__build_query()

    def __build_query(self):
        query_temp = ('SELECT %s %s FROM "SystemIndex" '
                      'WHERE WorkId IS NOT NULL %s %s %s %s')

        frags = ["", "", "", "", "", ""]
        frags[1] = '"{0}"'.format('", "'.join(self._select))
        frags[2] = "AND scope='file:%s'" % self._scope
        if self._limit:
            frags[0] = "TOP %d" % self._limit
        if self._sort_col:
            frags[5] = 'ORDER BY {0} {1}'.format(self._sort_col,
                                                 self._sort_order)
        if self._fpattern:
            if "%" in self._fpattern or "_" in self._fpattern:
                frags[3] = "AND System.FileName Like '%s'" % self._fpattern
            # if there are no wildcards we can use a contains which
            # is much faster as it uses the index
            else:
                frags[3] = ("AND Contains(System.FileName, "
                            "'{}')").format(self._fpattern)
        if self._exts:  # if we have file extensions
            # then we add a constraint against the System.ItemType column
            # in the form of Contains(System.ItemType, '.txt OR .doc OR .ppt')
            frags[4] = ("AND Contains(System.ItemType, "
                        ''''"{}"')''').format('" OR "'.join(self._exts))

        # print query_temp % tuple(frags)
        return query_temp % tuple(frags)

    def execute(self, connection=None):
        conn = connection or OleDb.connect(CONN_STRING)
        cursor = execute(str(self), conn)
        if connection is None:
            return cursor
            conn.close()
        return cursor

    def select(self, *cols):
        cols = list(cols)
        for index, col in enumerate(cols):
            cols[index] = SYSTEMINDEX_COLS.get(col, col)
        self._select = set(getattr(self, "_select", []))
        self._select = self._select.union(set(cols))
        self._select = list(self._select)

    def scope(self, path):
        # to add: test path if valid
        if not path:
            self._scope = ""
        self._scope = path
        return self

    def limit(self, n):
        self._limit = int(n)
        return self

    def sortby(self, col):
        if col is None:
            self._sort_col = None
        col = SYSTEMINDEX_COLS.get(col, col)
        self._sort_col = col
        return self

    def asc(self):
        self._sort_order = "ASC"
        return self

    def desc(self):
        self._sort_order = "DESC"
        return self

    def extensions(self, *exts):
        # ´´exts´´ without a dot at the beginning
        self._exts = set(getattr(self, "_exts", set()))
        self._exts = self._exts.union(set(exts))
        self._exts = list(self._exts)
        return self

    def fpattern(self, fpattern):
        # mapping cmd line style wildcards to SQL style wildcards
        for a, b in [("*", "%"), ("?", "_")]:
            fpattern = fpattern.replace(a, b)
        self._fpattern = fpattern
        return self


def execute(query, connection):
    cursor = connection.cursor()
    cursor.execute(str(query))
    return cursor


def itemurl2pathname(url):
    url = urllib2.splittype(url)[1]
    return urllib2.url2pathname(url)


def is_windows_search_avaible(scope):
    try:
        query = Query(limit=1, scope=scope)
        results = query.execute()
        results = results.rowcount
        return bool(results)
    except:  # add spec. Exceptions (!)
        return False


def test():
    import time
    start = time.time()

    # five most recently edited documents
    query = Query(select=["path"])
    query.limit(5).sortby("date").desc()
    query.extensions("pdf", "doc", "docx", "odt")

    cursor = query.execute()
    print "[*] Last five docs:"
    for row in cursor:
        print "\t", row[0]

    # memory usage of all indexed jpegs
    query = Query(select=["size"])
    query.extensions("jpg", "jpeg")

    cursor = query.execute()
    results = cursor.fetchall()

    total_size = sum(map(lambda r: r[0], results))
    print "[*] Total memory usage of all indexed jpegs: %d MBytes" % (
        total_size // 0x100000)

    # five most recently edited .py's
    query = Query(select=["url"])
    query.limit(5).sortby("date").desc()
    query.extensions("py")

    cursor = query.execute()
    print "[*] Last five pys:"
    for row in cursor:
        print "\t", itemurl2pathname(row[0])

    delta = time.time() - start
    print "[Finished in %.1fs]" % delta


if __name__ == "__main__":
    test()
