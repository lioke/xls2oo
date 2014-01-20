# -*- coding: utf8 -*-
__author__ = 'Alessandro Canova'
__version__ = 1.3
__date__ = "14/01/2014"
"""
This module is a little library to fast handling data table formatted in xls file

Changelog 1.3 - 14/01/2014
- filter mode on rows
- ctype on cells
Changelog 1.2 - 10/01/2014
- some litle improvement
Changelog 1.1 - 22/04/2013
- filter method on tables
"""

# XL_CELL_EMPTY     0 	empty string u''
# XL_CELL_TEXT 	    1 	a Unicode string
# XL_CELL_NUMBER 	2 	float
# XL_CELL_DATE 	    3 	float
# XL_CELL_BOOLEAN 	4 	int; 1 means TRUE, 0 means FALSE
# XL_CELL_ERROR 	5 	int representing internal Excel codes; for a text representation, refer to the supplied dictionary error_text_from_code
# XL_CELL_BLANK 	6 	empty string u''. Note: this type will appear only when open_workbook(..., formatting_info=True) is used.

xls2oo_DEBUG = False
import xlrd
import xlwt
import copy

XLS_TYPE_NONE = 0
XLS_TYPE_DB = 1
XLS_TYPE_TABLE = 2
XLS_TYPE_ROW = 3
XLS_TYPE_CELL = 4

XLS_EXTRA_ATTRIBUTE = ('id', 'pk')


class NotYetImplemented(Exception):
    def __init__(self, string_error):
        self.string_error = string_error
    def __str__(self):
        return repr(self.string_error)


class NotUniqueValue(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr("MultipleObjectsReturned: %d" % self.value)


class DoesNotExists(Exception):
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr("Attribute not defined: %s" % self.value)


class rawObj(object):
    """
    raw Object to store hidden informations
    """
    def __init__(self):
        """
        Set some important informations
        otype   =>  type of object None, table
        raw_name    => raw name object
        normalized_object_names    => ordered list of normalized name object contained (so table => columns name)
        raw_objects_names    =>  ordered list of raw name object contained (so table => columns name)
        """
        self.otype = XLS_TYPE_NONE
        self.raw_name = ""
        self.head_on_first_object = True
        self.normalized_object_names = []
        self.raw_objects_names =[]
        self.pk = -1


class XLSObject(object):
    """
    XLS generic object collections class (db, table, row)
    """
    def __init__(self, objects_list, raw_data):
        setattr(self, '__objects__', objects_list)
        setattr(self, '__raw__', raw_data)

    def __len__(self): return len(getattr(self, '__objects__'))

    def count(self):
        return self.__len__()

    def __getitem__(self, key):
        """
        Slice implementation
        """
        raw = getattr(self, '__raw__')
        objects = getattr(self, '__objects__')
        ret_obj = []
        if isinstance( key, slice ) :
            #Get the start, stop, and step from the slice
            ret_obj = [objects[ii] for ii in xrange(*key.indices(len(objects)))]
            return XLSObject(ret_obj, raw)
        elif isinstance( key, int ) :
            if key < 0 : #Handle negative indices
                key += len(objects)
            if key >= len(objects) :
                raise IndexError, "The index (%d) is out of range."%key
            #ret_obj.append(self.get_object_index(key))
            return self.get_object_index(key)#Get the data from elsewhere
        else:
            raise TypeError, "Invalid argument type."

    def get_object_index(self, index):
        if index > -1 and index < self.count():
            return getattr(self, '__objects__')[index]
        else:
            raise IndexError, "The index (%d) is out of range."%index

    def __iter__(self):
        for obj in getattr(self, '__objects__'):
            yield obj

    def all(self):
        """
        This method return all objects (for db, tables and row)
        """
        raw = getattr(self, '__raw__')
        if raw.otype==XLS_TYPE_DB:
            return XLS2OO(list_objects=getattr(self, '__objects__'), raw_data=raw)
        if raw.otype==XLS_TYPE_TABLE:
            return XLSObject(getattr(self, '__objects__'), raw)
        if raw.otype==XLS_TYPE_ROW:
            return XLSObject(getattr(self, '__objects__'), raw)
        raise NotYetImplemented('All method supported only for db, table and row')

    def get_or_create(self, **filter):
        """
        This method return a object that comply filters
        If object is not unique launch a exception, if objects don't exists create it
        """
        pass

    def get(self, **filters):
        """
        This method return a object that comply filters
        If object is not unique launch a exception
        """
        filtered = self.filter(**filters)
        raw = getattr(self, '__raw__')
        if raw.otype == XLS_TYPE_DB or raw.otype == XLS_TYPE_TABLE:
            if len(filtered) == 1:
                objects = getattr(filtered, '__objects__')
                return objects[0]
            else:
                raise NotUniqueValue(len(filtered))
        #elif raw.otype == XLS_TYPE_TABLE:
        #    raise NotYetImplemented('Not yet implemented (get table row). Sorry')

    def __filter__(self, filter_mode, filters):
        """
        This method return datalist filtered by filters
        If object is a Database:
            filters attributes: normalized table_name

        @filter_mode    =>  Select filter or exclude mode (if True filter mode, exclude_mode otherwise)
        """
        raw = getattr(self, '__raw__')

        if raw.otype not in (XLS_TYPE_DB, XLS_TYPE_TABLE):
            raise NotYetImplemented('Only database and table are supported for filtering')
        ret_filter_obj = []
        ret_exclude_obj = []
        #database filtering
        if raw.otype==XLS_TYPE_DB:
            if filter_mode == False:
                raise NotYetImplemented('Exclude mode for database is not yet implemented. Sorry!')
            if len(filters) != 1:
                raise NotYetImplemented('Only table_name=.. simple filter is implemented')
            indice = 0
            for table_name in raw.normalized_object_names:
                exclude = True
                for filter in filters:
                    if filter == 'table_name':
                        if table_name == filters[filter]:
                            exclude = False
                    else:
                        raise NotYetImplemented('Only table_name=.. filter is implemented')
                if not exclude:
                    new_raw = copy.deepcopy(raw)
                    new_raw.normalized_object_names = new_raw.normalized_object_names[indice : indice+1]
                    new_raw.raw_objects_names = new_raw.raw_objects_names[indice : indice+1]
                    objects = getattr(self, '__objects__')
                    return XLS2OO(list_objects=objects[indice : indice+1], raw_data=new_raw)
                indice = indice + 1

        #table filtering
        elif raw.otype==XLS_TYPE_TABLE:
            objects = getattr(self, '__objects__') #rows
            for obj in objects:
                exclude = False
                for filter in filters:
                    filter_part = filter.split('__')
                    if filter_part[0] not in raw.normalized_object_names and filter_part[0] not in XLS_EXTRA_ATTRIBUTE:
                        #I generate an Exception
                        exclude = True
                        raise DoesNotExists(filter_part[0])
                    index = raw.normalized_object_names.index(filter_part[0])
                    sub_obj = getattr(obj.objects.get_object_index(index), 'value')
                    #UNDONE: lookup implementation - here filter_part[1:n] can be a lookup to another table
                    if len(filter_part) == 1:
                        if sub_obj != filters[filter]:
                            exclude = True
                    else:
                        for f in filter_part[1:]:
                            if f == "exact":
                                if sub_obj != filters[filter]:
                                    exclude = True
                                    break
                            elif f == "iexact":
                                if sub_obj.lower() != filters[filter].lower():
                                    exclude = True
                                    break
                            elif f == "contains":
                                if sub_obj.find(filters[filter]) < 0:
                                    exclude = True
                                    break
                            elif f == "icontains":
                                if sub_obj.lower().find(filters[filter].lower()) < 0:
                                    exclude = True
                                    break
                            elif f == "startswith":
                                if not sub_obj.startswith(filters[filter]):
                                    exclude = True
                                    break
                            elif f == "istartswith":
                                if not sub_obj.lower().startswith(filters[filter].lower()):
                                    exclude = True
                                    break
                            elif f == "endswith":
                                if not sub_obj.endswith(filters[filter]):
                                    exclude = True
                                    break
                            elif f == "iendswith":
                                if not sub_obj.lower().endswith(filters[filter].lower()):
                                    exclude = True
                                    break
                            elif f == "lt":
                                if not sub_obj < filters[filter]:
                                    exclude = True
                                    break
                            elif f == "le":
                                if not sub_obj <= filters[filter]:
                                    exclude = True
                                    break
                            elif f == "gt":
                                if not sub_obj > filters[filter]:
                                    exclude = True
                                    break
                            elif f == "ge":
                                if not sub_obj >= filters[filter]:
                                    exclude = True
                                    break
                            elif f == "in":
                                if not sub_obj in filters[filter]:
                                    exclude = True
                                    break
                    if exclude:
                        break
                if not exclude:
                    ret_filter_obj.append(obj)
                else:
                    ret_exclude_obj.append(obj)
            #return XLSTable("filtered", ret_obj, raw)
            if filter_mode:
                return XLSObject(ret_filter_obj, raw)
            else:
                return XLSObject(ret_exclude_obj, raw)

        raise NotYetImplemented('Nothing filtered')

    def filter(self, **filters):
        return self.__filter__(True, filters)

    def exclude(self, **filters):
        return self.__filter__(False, filters)

    def order_by(self, **orders):
        for order in orders:
            reverse = False
            if order.startswith('-'):
                reverse = True
                order = order[1:]
        #UNDONE: not yet implemented
        raise  NotYetImplemented("Order by not yet implemented")
        #return self

    def reverse(self):
        """
        Reverse order object
        """
        raw = getattr(self, '__raw__')
        objects = getattr(self, '__objects__')
        ret_obj = []
        for o in reversed(objects):
            ret_obj.append(o)
        return XLSObject(ret_obj, raw)


class GenericXLSObject(object):
    """
    Generic table or row class
    """
    def __init__(self, object_list, raw_data_object, objects_raw_data_object):
        setattr(self, '__raw__', raw_data_object)
        self.objects = XLSObject(object_list, objects_raw_data_object)

    def __len__(self): return len(self.objects)

    def __getattr__(self, item):
        """
        This method dinamically define cell attributes on row
        """
        raw = super(GenericXLSObject, self).__getattribute__('__raw__')
        if item in ('id', 'pk'):
            if 'pk' in raw.normalized_object_names and 'id' not in raw.normalized_object_names:
                item = 'pk'
            if 'id' in raw.normalized_object_names and 'pk' not in raw.normalized_object_names:
                item = 'id'

            if 'id' not in raw.normalized_object_names and 'pk' not in raw.normalized_object_names:
                #UNDONE: Not yet implemented auto id
                raise NotYetImplemented('Auto id not yet implemented')
        if item in raw.normalized_object_names:
            index = raw.normalized_object_names.index(item)
            return self.objects.get_object_index(index)
        #elif item in ['id', 'pk']:
        #    return raw.pk
        else:
            super(GenericXLSObject, self).__getattr__(item)


class XLSCell(object):
    """
    XLS Cell class
    """
    def __init__(self, value, ctype=None, xfindex=None):
        self.__ctype__ = ctype
        self.value = value
        self.__xfindex__ = xfindex

    def __getattribute__(self, name):
        if name == 'ctype':
            return self.__ctype__
        return super(XLSCell, self).__getattribute__(name)


    def __repr__(self):
        try:
            return unicode(self.value)
        except:
            pass
        return super(XLSCell, self).__repr__()


class XLSRow(GenericXLSObject):
    """
    XLS Row class
    """
    def __init__(self, xls_cells, raw_attributes_object):
        super(XLSRow, self).__init__(xls_cells, raw_attributes_object, raw_attributes_object)

class XLSTable(GenericXLSObject):
    """

    """
    def __init__(self, normalized_tablename, xls_rows, raw_attributes_object=None):
        """
        Constructor
        This Method fill data structure with:
        @Normalized tablename   =>  Name of table normalized
        @Raw attribute object =>  A object of raw informations
            It contains follow attributes:
            dimension  key  =>  numbers of tables cols
            head_on_first_object    => if True first row of table contain the head
            normalized_object_names =>  a ordered list of normalized colnames
                head_on_first_object is True  =>  normalized colname is the first row cell normalized when applicable, otherwise is "d_#col"
                head_on_first_object is False  =>  "d_#col"
            raw_name    =>  table raw name
            raw_objects_names => a ordered list of colnames
                head_on_first_object is True  =>  colname is the first row cell when applicable, otherwise is "c_#col"
                head_on_first_object is False  =>  "c_#col"
        """
        super(XLSTable, self).__init__(xls_rows, raw_attributes_object, raw_attributes_object)
        self.name = normalized_tablename

    # def __get_col_number_by_raw_name__(self, rawname):
    #     """
    #     This method get raw attribute c_n in raw format and return n
    #     """
    #     if rawname.startswith('c_'):
    #         return int(rawname[2:])
    #     else:
    #         return None

    # def get_table_row_as_dict(self, rownum):
    #     """
    #     This method return a rownum row of a table
    #     If rownum table don't exists then it retrieve None
    #     Input:
    #         @rownum  => number of the row that have to be retrieved
    #     Output:
    #         A dictionary of data with virtual name column
    #     """
    #     try:
    #         row = self.rows[rownum]
    #         rowret = dict()
    #         for attribute in getattr(self, '__rawheadsorder__'):
    #             rowret[getattr(self, '__reverseheads__')[attribute]['virtualname']] = row.cells[self.__get_col_number_by_raw_name__(attribute)].value
    #     except:
    #         return None
    #     return rowret

    # def get_table_lenght(self):
    #     """
    #     This method return the number of record of a table
    #     If table don't exists then it retrieve zero
    #     Output:
    #         Number of records in the table if table exists otherwise zero
    #     """
    #     try:
    #         return len(self.rows)
    #     except:
    #         return 0


class XLS2OO(GenericXLSObject):
    """

    """
    def __init__(self, xls_file_path=None, head_on_first_row=True, encode=None, list_objects=(), raw_data=rawObj()):
        """
        Input:
            @xls_file_path    =>  path of the xls file
            @head_on_first_row => if True it assume that each table has a head on the first row
            @encode => Only for older xls format, it override the default encode format (utf16le)

            @list_objects    =>  a list of ordered table objects
            @raw_data   =>  db raw data info
        Define:
            objects => ordered list of table
            __raw__ => objects of raw informations such:
                xls_file_path =>  xls filepath origin (string)
                encode  =>  encode ovverride
        """
        if xls_file_path == None:
            self.xlsfilepath = raw_data.raw_name
            super(XLS2OO, self).__init__(list_objects, raw_data, raw_data)
        else:
            self.xlsfilepath = xls_file_path
            raw = rawObj()
            raw.encode = encode if encode else "utf16le"
            raw.head_on_first_object = head_on_first_row
            raw.xls_file_path = xls_file_path
            tables, raw_data = self.__load_xls_file__(raw.xls_file_path, raw.head_on_first_object, raw.encode)
            super(XLS2OO, self).__init__(tables, raw, raw_data)
        return

    def __normalize_name__(self, name):
        """
        This method normalize a name with these rules:
        A) Each letter is in lower case
        B) Convert to pure ascii string
        C) All multiple spaces are replaced by a single underscore
        """
        import unicodedata
        #A
        normalized_name = name.lower()
        #B
        #normalized_name = normalized_name.replace(u'è', 'e').replace(u'é', 'e').replace(u'à', 'a').replace(u'ò', 'o')
        try:
            normalized_name = unicode(normalized_name)
        except:
            pass
        normalize_name = unicodedata.normalize('NFKD', normalized_name).encode('ascii','ignore')
        #C
        while True:
            if normalized_name.find('  ') > -1:
                normalized_name = normalized_name.replace('  ', ' ')
            else:
                break
        normalized_name = normalized_name.replace(' ', '_')
        return normalized_name

    def __load_xls_file__(self, xls_file_path, head_on_first_row, encode):
        """
        This method open a xls file (xls_file_path) and load it data into a memory structure.
        Input:
            @xls_file_path    =>  full path of the xls file to open
            @head_on_first_row =>  if True first row of each sheet is the head of the table
            @encode =>  If encoding_override is setting to encode (only for older xls file format)
        Output:
            This function fill table objects
        """
        try:
            workbook = xlrd.open_workbook(xls_file_path, formatting_info=True, encoding_override=encode)
        except:
            print "Error to opening file: %s" % xls_file_path
            return False

        xls_tables = []

        tables = workbook.sheet_names()
        raw_db_info = rawObj()
        raw_db_info.otype = XLS_TYPE_DB
        raw_db_info.head_on_first_object = False
        raw_db_info.raw_name = xls_file_path
        raw_db_info.pk = 1

        for table in tables:
            normalized_table_name = self.__normalize_name__(table)
            raw_db_info.normalized_object_names.append(normalized_table_name)
            raw_db_info.raw_objects_names.append(table)

            raw_table_info = rawObj()
            raw_table_info.otype = XLS_TYPE_TABLE
            raw_table_info.head_on_first_object = head_on_first_row
            raw_table_info.raw_name = table
            raw_table_info.pk = len(raw_db_info.normalized_object_names)

            worksheet = workbook.sheet_by_name(table)
            first_data_record = True
            coltype = dict()
            table_rows = []
            for rownum in range(worksheet.nrows):
                row = worksheet.row(rownum)
                if rownum == 0:
                    for ncol in range(0, len(row)):
                        if head_on_first_row:
                            #heads are on the first row
                            realname = row[ncol].value
                            if not realname:
                                realname = "__virtual_%d__" % ncol
                            virtualname = self.__normalize_name__(realname)
                        else:
                              realname = "__virtual_%d__" % ncol
                              virtualname = "d_%d" % ncol

                        raw_table_info.raw_objects_names.append(realname)
                        raw_table_info.normalized_object_names.append(virtualname)
                if head_on_first_row and rownum==0:
                    continue
                #now populate the table
                row_cells = []
                for ncol in range(0, len(row)):
                    celltype = row[ncol].ctype
                    cellvalue = row[ncol].value
                    if(celltype == xlrd.XL_CELL_NUMBER):
                        if int(cellvalue) == cellvalue:
                            cellvalue = int(cellvalue)
                    cellxfindex = row[ncol].xf_index
                    row_cells.append(XLSCell(cellvalue, celltype, cellxfindex))
                    if(first_data_record):
                        coltype[ncol] = celltype
                    if(coltype[ncol] != celltype):
                        if xls2oo_DEBUG:
                            #UNDONE: print alert because all data in the same column aren't of the same type
                            print "Alert, data on the column %d are a inchoerence on row %d" % (ncol, rownum)

                raw_row_info = rawObj()
                raw_row_info.otype = XLS_TYPE_ROW
                raw_row_info.raw_name = str(rownum)
                raw_row_info.raw_objects_names = raw_table_info.raw_objects_names
                raw_row_info.normalized_object_names = raw_table_info.normalized_object_names
                raw_row_info.pk = rownum + 1 if head_on_first_row else rownum

                first_data_record = False
                table_rows.append(XLSRow(row_cells, raw_row_info))

            xls_tables.append(XLSTable(normalized_table_name, table_rows, raw_table_info))
        return xls_tables, raw_db_info

    def make_table_index(self, table, attribute):
        """
        This method make a index for the table based on column attrbute
        Attribute can be a number (number of the column)
        """
        #UNDONE: to be defined
        pass

    def get_table(self, tablename):
        """
        This method return a table object by it name
        If table exists return a table otherwise return an empty table
        Input:
            @tablename  => name to table that have to be retrieved
        """
        raw = getattr(self.objects, '__raw__')

        if tablename in raw.normalized_object_names:
            index = raw.normalized_object_names.index(tablename)
            objects = getattr(self.objects, '__objects__')
            return objects[index]
        else:
            return XLSTable()





