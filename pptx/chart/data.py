# encoding: utf-8

"""
ChartData and related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

import warnings

from ..oxml import parse_xml
from ..oxml.ns import nsdecls
from .xlsx import WorkbookWriter
from .xmlwriter import ChartXmlWriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name
from xml.sax.saxutils import escape


class ChartData(object):
    """
    Accumulates data specifying the categories and series values for a plot
    and acts as a proxy for the chart data table that will be written to an
    Excel worksheet. Used as a parameter in :meth:`shapes.add_chart` and
    :meth:`Chart.replace_data`.
    """
    def __init__(self):
        super(ChartData, self).__init__()
        self._categories = []
        self._series_lst = []

    def add_series(self, name, values):
        """
        Add a series to this data set entitled *name* and the data points
        specified by *values*, an iterable of numeric values.
        """
        series_idx = len(self._series_lst)
        series = _SeriesData(series_idx, name, values, self._categories)
        self._series_lst.append(series)

    @property
    def categories(self):
        """
        Read-write. The sequence of category label strings to use in the
        chart. Any type that is iterable over a sequence of strings can be
        assigned, e.g. a list, tuple, or iterator.
        """
        return tuple(self._categories)

    @categories.setter
    def categories(self, categories):
        # Contents need to be replaced in-place so reference sent to
        # _SeriesData objects retain access to latest values
        self._categories[:] = categories

    @property
    def series(self):
        """
        An snapshot of the current |_SeriesData| objects in this chart data
        contained in an immutable sequence.
        """
        return tuple(self._series_lst)

    @property
    def xlsx_blob(self):
        """
        Return a blob containing an Excel workbook file populated with the
        categories and series in this chart data object.
        """
        return WorkbookWriter.xlsx_blob(self.categories, self._series_lst)

    def xml_bytes(self, chart_type):
        """
        Return a blob containing the XML for a chart of *chart_type*
        containing the series in this chart data object, as bytes suitable
        for writing directly to a file.
        """
        return self._xml(chart_type).encode('utf-8')

    def _xml(self, chart_type):
        """
        Return (as unicode text) the XML for a chart of *chart_type* having
        the categories and series in this chart data object. The XML is
        a complete XML document, including an XML declaration specifying
        UTF-8 encoding.
        """
        return ChartXmlWriter(chart_type, self._series_lst).xml


class _SeriesData(object):
    """
    Like |ChartData|, a data transfer object, but specific to the data
    specifying a series. In addition, this object also provides XML
    generation for the ``<c:ser>`` element subtree.
    """
    def __init__(self, series_idx, name, values, categories):
        super(_SeriesData, self).__init__()
        self._series_idx = series_idx
        self._name = name
        self._values = values
        self._categories = categories

    def __len__(self):
        """
        The number of values this series contains.
        """
        return len(self._values)

    @property
    def cat(self):
        """
        The ``<c:cat>`` element XML for this series, as an oxml element.
        """
        xml = self._cat_tmpl.format(
            wksht_ref=self._categories_ref, cat_count=len(self._categories),
            cat_pt_xml=self._cat_pt_xml, nsdecls=' %s' % nsdecls('c')
        )
        return parse_xml(xml)

    @property
    def cat_xml(self):
        """
        The unicode XML snippet for the ``<c:cat>`` element for this series,
        containing the category labels and spreadsheet reference.
        """
        return self._cat_tmpl.format(
            wksht_ref=self._categories_ref, cat_count=len(self._categories),
            cat_pt_xml=self._cat_pt_xml, nsdecls=''
        )

    @property
    def index(self):
        """
        The zero-based index of this series within the plot.
        """
        return self._series_idx

    @property
    def name(self):
        """
        The name of this series.
        """
        return self._name

    @property
    def tx(self):
        """
        Return a ``<c:tx>`` oxml element for this series, containing the
        series name.
        """
        xml = self._tx_tmpl.format(
            wksht_ref=self._series_name_ref, series_name=escape(self.name),
            nsdecls=' %s' % nsdecls('c')
        )
        return parse_xml(xml)

    @property
    def tx_xml(self):
        """
        Return the ``<c:tx>`` element for this series as unicode text. This
        element contains the series name.
        """
        return self._tx_tmpl.format(
            wksht_ref=self._series_name_ref, series_name=escape(self.name),
            nsdecls=''
        )

    @property
    def val(self):
        """
        The ``<c:val>`` XML for this series, as an oxml element.
        """
        xml = self._val_tmpl.format(
            wksht_ref=self._values_ref, val_count=len(self),
            val_pt_xml=self._val_pt_xml, nsdecls=' %s' % nsdecls('c')
        )
        return parse_xml(xml)

    @property
    def val_xml(self):
        """
        Return the unicode XML snippet for the ``<c:val>`` element describing
        this series.
        """
        return self._val_tmpl.format(
            wksht_ref=self._values_ref, val_count=len(self),
            val_pt_xml=self._val_pt_xml, nsdecls=''
        )

    @property
    def values(self):
        """
        The values in this series as a sequence of float.
        """
        return self._values

    @property
    def _categories_ref(self):
        """
        The Excel worksheet reference to the categories for this series.
        """
        end_row_number = len(self._categories) + 1
        return "Sheet1!$A$2:$A$%d" % end_row_number

    @property
    def _cat_pt_xml(self):
        """
        The unicode XML snippet for the ``<c:pt>`` elements containing the
        category names for this series.
        """
        xml = ''
        for idx, name in enumerate(self._categories):
            xml += (
                '                <c:pt idx="%d">\n'
                '                  <c:v>%s</c:v>\n'
                '                </c:pt>\n'
            ) % (idx, name)
        return xml

    @property
    def _cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series, containing
        the category labels and spreadsheet reference.
        """
        return (
            '          <c:cat{nsdecls}>\n'
            '            <c:strRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:strCache>\n'
            '                <c:ptCount val="{cat_count}"/>\n'
            '{cat_pt_xml}'
            '              </c:strCache>\n'
            '            </c:strRef>\n'
            '          </c:cat>\n'
        )

    @property
    def _col_letter(self):
        """
        The letter of the Excel worksheet column in which the data for this
        series appears.
        """
        return chr(ord('B') + self._series_idx)

    @property
    def _series_name_ref(self):
        """
        The Excel worksheet reference to the name for this series.
        """
        return "Sheet1!$%s$1" % self._col_letter

    @property
    def _tx_tmpl(self):
        """
        The string formatting template for the ``<c:tx>`` element for this
        series, containing the series title and spreadsheet range reference.
        """
        return (
            '          <c:tx{nsdecls}>\n'
            '            <c:strRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:strCache>\n'
            '                <c:ptCount val="1"/>\n'
            '                <c:pt idx="0">\n'
            '                  <c:v>{series_name}</c:v>\n'
            '                </c:pt>\n'
            '              </c:strCache>\n'
            '            </c:strRef>\n'
            '          </c:tx>\n'
        )

    @property
    def _val_pt_xml(self):
        """
        The unicode XML snippet containing the ``<c:pt>`` elements for this
        series.
        """
        xml = ''
        for idx, value in enumerate(self._values):
            xml += (
                '                <c:pt idx="%d">\n'
                '                  <c:v>%s</c:v>\n'
                '                </c:pt>\n'
            ) % (idx, value)
        return xml

    @property
    def _val_tmpl(self):
        """
        The template for the ``<c:val>`` element for this series, containing
        the series values and their spreadsheet range reference.
        """
        return (
            '          <c:val{nsdecls}>\n'
            '            <c:numRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:numCache>\n'
            '                <c:formatCode>General</c:formatCode>\n'
            '                <c:ptCount val="{val_count}"/>\n'
            '{val_pt_xml}'
            '              </c:numCache>\n'
            '            </c:numRef>\n'
            '          </c:val>\n'
        )

    @property
    def _values_ref(self):
        """
        The Excel worksheet reference to the values for this series (not
        including the series name).
        """
        return "Sheet1!$%s$2:$%s$%d" % (
            self._col_letter, self._col_letter, len(self._values)+1
        )


class ChartDataMoreDetails(ChartData):
    """
    Subclass of ChartData, support categories and vals with more details:
    categories with multiple levels, categories and vals with blanks.

    See also :class: `ChartData`.
    """
    def __init__(self):
        super(ChartDataMoreDetails, self).__init__()
        self._categories_len = None
        self._values_len = None

    @property
    def categories_len(self):
        """
        Read-write. The length of categories. Assigned value
        will be applied to all sers
        """
        return self._categories_len

    @categories_len.setter
    def categories_len(self, value):
        self._categories_len = value
        #make sure all sers have this categories_len
        for series in self._series_lst:
            series.categories_len = value

    @property
    def values_len(self):
        """
        Read-write. The length of values.Assigned value
        will be applied to all sers
        """
        return self._values_len

    @values_len.setter
    def values_len(self, value):
        self._values_len = value
        if self._values_len < self._categories_len:
            warnings.warn(
                '''Length of values is less than that of categories.
Over bound categories will not be displayed.''')
        #make sure all sers have this values_len
        for series in self._series_lst:
            series.values_len = value


    def add_series(self, name, values, values_len=None, format_code=None):
        """
        Add a series to this data set entitled *name* and the data points
        specified by *values*, an iterable of numeric values.
        """
        series_idx = len(self._series_lst)
        series = _SeriesDataMoreDetails(
            series_idx, name, values,
            self.categories,
            values_len = values_len or self._values_len,
            categories_len=self._categories_len,
            format_code=format_code,
            )
        self._series_lst.append(series)


class _SeriesDataMoreDetails(_SeriesData):
    """
    Subclass of _SeriesData, support categories and vals with more details:

     * categories with multiple levels
     * categories and vals with blanks
     * column letters exceeding 'Z'
     * formatCode

    Arguments : values & categories must be 2D sequence of (idx, value)

    See also : :class: `_SeriesData`.
    """
    def __init__(self, series_idx, name, values, categories,
                 values_len=None, categories_len=None, format_code=None):
        super(_SeriesDataMoreDetails, self).__init__(series_idx, name,
                                                     values, categories)
        self._values_len = values_len or max(i[0] for i in values) + 1
        self._auto_values_len = True if values_len else False
        if categories_len is None and (not self._categories is None) and (
                len(self._categories) != 0):
            self._categories_len = max(max(j[0] for j in i)
                                for i in self._categories) + 1
        else:
            self._categories_len = categories_len
        if self._values_len != self._categories_len:
            warnings.warn('''Categories and Values have different lengths.
 Will break data range adjustment by dragging in MS PowerPoint.''')
        self._format_code = format_code or 'General'

    @property
    def categories_len(self):
        """
        Read-write. The length of categories.
        """
        return self._categories_len

    @categories_len.setter
    def categories_len(self, value):
        self._categories_len = value

    @property
    def values_len(self):
        """
        Read-write. The length of values.
        """
        return self._values_len

    @values_len.setter
    def values_len(self, value):
        self._values_len = value

    @property
    def format_code(self):
        """
        format code string in ``<c:formatCode>`` element
        """
        return self._format_code

    @property
    def is_cat_multilvl(self):
        """
        whether ``<c:cat>`` element has multiple levels
        """
        if len(self._categories) > 1:
            return True
        else:
            return False

    @property
    def prefix(self):
        """
        prefix for ``<c:*Ref>`` and ``<c:*Cache>`` element
        """
        if self.is_cat_multilvl:
            return 'multiLvlStr'
        else:
            return 'str'

    @property
    def cat(self):
        """
        The ``<c:cat>`` element XML for this series, as an oxml element.
        """
        if self._categories_len > 0:
            xml = self._cat_tmpl.format(
                prefix=self.prefix,
                wksht_ref=self._categories_ref, cat_count=self._categories_len,
                cat_pt_xml=self._cat_pt_xml, nsdecls=' %s' % nsdecls('c')
            )
            return parse_xml(xml)
        else:
            return None


    @property
    def cat_xml(self):
        """
        The unicode XML snippet for the ``<c:cat>`` element for this series,
        containing the category labels and spreadsheet reference.
        """
        if self._categories_len > 0:
            return self._cat_tmpl.format(
                prefix=self.prefix,
                wksht_ref=self._categories_ref, cat_count=self._categories_len,
                cat_pt_xml=self._cat_pt_xml, nsdecls=''
            )
        else:
            return ''

    @property
    def val(self):
        """
        The ``<c:val>`` XML for this series, as an oxml element.
        """
        xml = self._val_tmpl.format(
            wksht_ref=self._values_ref, val_count=self._values_len,
            format_code=self._format_code,
            val_pt_xml=self._val_pt_xml, nsdecls=' %s' % nsdecls('c')
        )
        return parse_xml(xml)

    @property
    def val_xml(self):
        """
        Return the unicode XML snippet for the ``<c:val>`` element describing
        this series.
        """
        return self._val_tmpl.format(
            wksht_ref=self._values_ref, val_count=self._values_len,
            format_code=self._format_code,
            val_pt_xml=self._val_pt_xml, nsdecls=''
        )

    @property
    def values(self):
        """
        The values in this series as a tuple of a sequence of float.
        """
        return self._values

    @property
    def _categories_ref(self):
        """
        The Excel worksheet reference to the categories for this series.
        """
        end_col_number = len(self._categories) - 1
        end_row_number = self._categories_len
        return "Sheet1!$A$2:%s" % xl_rowcol_to_cell(
            end_row_number, end_col_number,
            row_abs=True, col_abs=True)

    @property
    def _cat_pt_xml(self):
        """
        The unicode XML snippet for the ``<c:pt>`` elements containing the
        category names for this series.
        """
        xml = ''
        if self.is_cat_multilvl:
            lvl_start_tag = '                  <c:lvl>\n'
            lvl_end_tag = '                  </c:lvl>\n'
            pt_indent_spaces = '  '
        else:
            lvl_start_tag = ''
            lvl_end_tag = ''
            pt_indent_spaces = ''
        pt_xml = pt_indent_spaces.join(
            ('',
             '                  <c:pt idx="%d">\n',
             '                    <c:v>%s</c:v>\n',
             '                  </c:pt>\n',))
        #ref lvl is in reverse sequence in xml
        loop_range = range(len(self._categories))
        loop_range.reverse()
        for ilvl in loop_range:
            lvl = self._categories[ilvl]
            xml += lvl_start_tag
            for idx, name in lvl:
                #ignore idx out bound
                if idx < self.categories_len:
                    xml += pt_xml % (idx, name)
            xml += lvl_end_tag
        return xml

    @property
    def _cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series, containing
        the category labels and spreadsheet reference.
        """
        return (
            '          <c:cat{nsdecls}>\n'
            '            <c:{prefix}Ref>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:{prefix}Cache>\n'
            '                <c:ptCount val="{cat_count}"/>\n'
            '{cat_pt_xml}'
            '              </c:{prefix}Cache>\n'
            '            </c:{prefix}Ref>\n'
            '          </c:cat>\n'
        )

    @property
    def _col_letter(self):
        """
        The letter of the Excel worksheet column in which the data for this
        series appears.
        """
        return xl_col_to_name(max(1, len(self._categories)) + self._series_idx)

    @property
    def _val_pt_xml(self):
        """
        The unicode XML snippet containing the ``<c:pt>`` elements for this
        series.
        """
        xml = ''
        for idx, value in self._values:
            if idx < self.values_len:
                xml += (
                    '                <c:pt idx="%d">\n'
                    '                  <c:v>%s</c:v>\n'
                    '                </c:pt>\n'
                    ) % (idx, value)
        return xml

    @property
    def _val_tmpl(self):
        """
        The template for the ``<c:val>`` element for this series, containing
        the series values and their spreadsheet range reference.
        """
        return (
            '          <c:val{nsdecls}>\n'
            '            <c:numRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:numCache>\n'
            '                <c:formatCode>{format_code}</c:formatCode>\n'
            '                <c:ptCount val="{val_count}"/>\n'
            '{val_pt_xml}'
            '              </c:numCache>\n'
            '            </c:numRef>\n'
            '          </c:val>\n'
        )

    @property
    def _values_ref(self):
        """
        The Excel worksheet reference to the values for this series (not
        including the series name).
        """
        return "Sheet1!${col_letter}$2:${col_letter}${end_row_number}".format(
            col_letter = self._col_letter,
            end_row_number = self._values_len + 1,
        )

    @property
    def name(self):
        """
        The name of this series.
        """
        return self._name

    @name.setter
    def name(self, value):
        #name setter
        self._name = value

    @property
    def values(self):
        """
        The values in this series as a sequence of float.
        """
        return self._values

    @values.setter
    def values(self, _values):
        #values setter
        self._values = _values
        #update values len if auto
        if self._auto_values_len:
            self._values_len = max(i[0] for i in _values) + 1
