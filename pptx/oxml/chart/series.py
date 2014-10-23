# encoding: utf-8

"""
Series-related oxml objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from ..simpletypes import XsdUnsignedInt
from ..xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrOne, OneOrMore
)


class CT_SeriesComposite(BaseOxmlElement):
    """
    ``<c:ser>`` custom element class. Note there are several different series
    element types in the schema, such as ``CT_LineSer`` and ``CT_BarSer``,
    but they all share the same tag name. This class acts as a composite and
    depends on the caller not to do anything invalid for a series belonging
    to a particular plot type.
    """
    idx = OneAndOnlyOne('c:idx')
    order = OneAndOnlyOne('c:order')
    tx = ZeroOrOne('c:tx')      # provide override for _insert_tx()
    spPr = ZeroOrOne('c:spPr')  # provide override for _insert_spPr()
    invertIfNegative = ZeroOrOne('c:invertIfNegative')  # provide _insert..()
    cat = ZeroOrOne('c:cat', successors=(
        'c:val', 'c:smooth', 'c:shape', 'c:extLst'
    ))
    val = ZeroOrOne('c:val', successors=('c:smooth', 'c:shape', 'c:extLst'))
    smooth = ZeroOrOne('c:smooth', successors=('c:extLst',))

    @property
    def val_pts(self):
        """
        The sequence of ``<c:pt>`` elements under the ``<c:val>`` child
        element, ordered by the value of their ``idx`` attribute.
        """
        val_pts = self.xpath('./c:val//c:pt')
        return sorted(val_pts, key=lambda pt: pt.idx)

    def _insert_invertIfNegative(self, invertIfNegative):
        """
        invertIfNegative has a lot of successors and they vary depending on
        the series type, so easier just to insert it "manually" as it's close
        to a required element.
        """
        if self.spPr is not None:
            self.spPr.addnext(invertIfNegative)
        elif self.tx is not None:
            self.tx.addnext(invertIfNegative)
        else:
            self.order.addnext(invertIfNegative)
        return invertIfNegative

    def _insert_spPr(self, spPr):
        """
        spPr has a lot of successors and it varies depending on the series
        type, so easier just to insert it "manually" as it's close to a
        required element.
        """
        if self.tx is not None:
            self.tx.addnext(spPr)
        else:
            self.order.addnext(spPr)
        return spPr

    def _insert_tx(self, tx):
        self.order.addnext(tx)
        return tx


class CT_StrVal_NumVal_Composite(BaseOxmlElement):
    """
    ``<c:pt>`` element, can be either CT_StrVal or CT_NumVal complex type.
    Using this class for both, differentiating as needed.
    """
    v = OneAndOnlyOne('c:v')
    idx = RequiredAttribute('idx', XsdUnsignedInt)

    @property
    def value(self):
        """
        The float value of the text in the required ``<c:v>`` child.
        """
        return float(self.v.text)

class _Base_Seq(BaseOxmlElement):
    """
    base class for sequence element
    provides similar properties and methods for element with ref and cache
    """
    #ref element must be implemented in subclass

    @property
    def ref(self):
        """
        ref element, must be implemented in subclass
        """
        raise NotImplementedError(
            'property ref must be implemented in subclass')

    @property
    def text_ref(self):
        """
        reference text
        """
        return self.ref.text_cf

    @property
    def val_ptCount(self):
        """
        val of ptCount
        """
        return self.ref.cache.ptCount.val

    @property
    def tuples_pts(self):
        """
        tuples of pts
        """
        return self.ref.cache.tuple_pts


class CT_Val(_Base_Seq):
    """
    ``<c:val>`` element, contains values
    """
    _ref = ZeroOrOne('c:numRef')

    @property
    def ref(self):
        """
        ref element of ``<c:numRef>`` child
        """
        return self._ref


class CT_Cat(_Base_Seq):
    """
    ``<c:cat>`` element, contains categories
    """
    _multilvlstrref = ZeroOrOne('c:multiLvlStrRef')
    _strref = ZeroOrOne('c:strRef')

    @property
    def is_multilvl(self):
        """
        True if with ``c:multiLvlStrRef>`` child
        """
        if self._multilvlstrref is None:
            return False
        else:
            return True

    @property
    def ref(self):
        """
        ref element of ``<c:multiLvlStrRef>`` or ``<c:strRef>`` child
        """
        if self.is_multilvl:
            return self._multilvlstrref
        else:
            return self._strref

    @property
    def tuples_pts(self):
        """
        tuples of pts
        """
        if self.is_multilvl:
            return tuple(
                (lvl.tuple_pts for lvl in self.ref.cache.lvl_lst)
                )
        else:
            return (self.ref.cache.tuple_pts,)


class _Base_Ref(BaseOxmlElement):
    """
    base class for ``<c:multiLvlStrRef>``, ``<c:strRef>``, and ``<c:numCache>``
    element
    """
    cf = ZeroOrOne('c:f', successors=('c:multiLvlStrCache', 'c:strCache',
                                      'c:numCache'))

    @property
    def text_cf(self):
        """
        reference text of ``<c:f>`` child.
        """
        if not self.cf is None:
            return self.cf.text


class CT_MultiLvlStrRef(_Base_Ref):
    """
    ``<c:multiLvlStrRef>`` element
    """
    cache = OneAndOnlyOne('c:multiLvlStrCache')


class CT_StrRef(_Base_Ref):
    """
    ``<c:strRef>`` element
    """
    cache = OneAndOnlyOne('c:strCache')

class CT_NumRef(_Base_Ref):
    """
    ``<c:strRef>`` element
    """
    cache = OneAndOnlyOne('c:numCache')


class _Base_Cache(BaseOxmlElement):
    """
    base class for  ``<c:multiLvlStrCache>``, ``<c:strCache>``, and ``<c:numCache>`` element
    """
    formatCode = ZeroOrOne('c:formatCode', successors=('c:ptCount',))
    ptCount = OneAndOnlyOne('c:ptCount')


class CT_MultiLvlStrCache(_Base_Cache):
    """
    ``<c:multiLvlStrCache>`` element
    """
    lvl = OneOrMore('c:lvl')


class _PtMixin(BaseOxmlElement):
    """
    mixin class for  ``<c:pt>`` children, provides pt_tuples
    """
    pt = OneOrMore('c:pt')

    @property
    def tuple_pts(self):
        """
        return pt tuples as : (idx, v.text)
        """
        return tuple(
            ((pt.idx, pt.v.text) for pt in self.pt_lst)
            )


class CT_Lvl(_PtMixin):
    """
    ``<c:lvl>`` element
    """
    pass


class CT_StrCache(_Base_Cache, _PtMixin):
    """
    ``<c:strCache>`` element
    """
    pass


class CT_NumCache(_Base_Cache, _PtMixin):
    """
    ``<c:numCache>`` element
    """
    @property
    def text_fomatCode(self):
        """
        return text of formatCode
        """
        return self.formatCode.text

    @property
    def tuple_pts(self):
        """
        try return pt tuple as : (idx, float(v.text))
        and return (idx, v.text) if failed
        """
        pt_tuple_lst = list()
        for pt in self.pt_lst:
            try:
                value = float(pt.v.text)
            except ValueError:
                value = pt.v.text
            pt_tuple_lst.append((pt.idx, value))
        return tuple(pt_tuple_lst)
