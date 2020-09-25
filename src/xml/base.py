from lxml.etree import Element, SubElement, QName, tounicode


class NS_DEF_TYPE(object):
    wpc         = 'wpc'
    cx          = 'cx'
    cx1         = 'cx1'
    cx2         = 'cx2'
    cx3         = 'cx3'
    cx4         = 'cx4'
    cx5         = 'cx5'
    cx6         = 'cx6'
    cx7         = 'cx7'
    cx8         = 'cx8'
    mc          = 'mc'
    aink        = 'aink'
    am3d        = 'am3d'
    o           = 'o'
    r           = 'r'
    m           = 'm'
    v           = 'v'
    wp14        = 'wp14'
    wp          = 'wp'
    w10         = 'w10'
    w           = 'w' 
    w14         = 'w14'
    w15         = 'w15'
    w16cex      = 'w16cex'
    w16cid      = 'w16cid'
    w16         = 'w16'
    w16se       = 'w16se'
    wpg         = 'wpg'
    wpi         = 'wpi'
    wne         = 'wne'
    wps         = 'wps'


class DOCX_NS:
    @classmethod
    def ns(self, ns=NS_DEF_TYPE.w):
        return {
                        'wpc'       : "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
                        'cx'        : "http://schemas.microsoft.com/office/drawing/2014/chartex",
                        'cx1'       : "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
                        'cx2'       : "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
                        'cx3'       : "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
                        'cx4'       : "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
                        'cx5'       : "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
                        'cx6'       : "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
                        'cx7'       : "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
                        'cx8'       : "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
                        'mc'        : "http://schemas.openxmlformats.org/markup-compatibility/2006",
                        'aink'      : "http://schemas.microsoft.com/office/drawing/2016/ink",
                        'am3d'      : "http://schemas.microsoft.com/office/drawing/2017/model3d",
                        'o'         : "urn:schemas-microsoft-com:office:office",
                        'r'         : "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        'm'         : "http://schemas.openxmlformats.org/officeDocument/2006/math",
                        'v'         : "urn:schemas-microsoft-com:vml",
                        'wp14'      : "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
                        'wp'        : "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                        'w10'       : "urn:schemas-microsoft-com:office:word",
                        'w'         : "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                        'w14'       : "http://schemas.microsoft.com/office/word/2010/wordml",
                        'w15'       : "http://schemas.microsoft.com/office/word/2012/wordml",
                        'w16cex'    : "http://schemas.microsoft.com/office/word/2018/wordml/cex",
                        'w16cid'    : "http://schemas.microsoft.com/office/word/2016/wordml/cid",
                        'w16'       : "http://schemas.microsoft.com/office/word/2018/wordml",
                        'w16se'     : "http://schemas.microsoft.com/office/word/2015/wordml/symex",
                        'wpg'       : "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
                        'wpi'       : "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
                        'wne'       : "http://schemas.microsoft.com/office/word/2006/wordml",
                        'wps'       : "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
                }[ns]    

class DOCX_ELEMENT_FACTORY:
    def __init__(self):
        self.attribs    = []
        self.node       = None

    def create(self, nsdef=NS_DEF_TYPE.w, name='p'):
        self.node       = Element(QName(DOCX_NS.ns(nsdef), name), nsmap={nsdef:DOCX_NS.ns(nsdef)})
        return self
    
    def add_attribs(self, nsdef, name, value, qname=True):
        if qname:
            self.attribs.append({'qname': QName(DOCX_NS.ns(nsdef), name), 'val':value})
        else:
            self.attribs.append({'qname': name, 'val':value})
    
    def get(self):
        if len(self.attribs) > 0:
            attrib = {}
            for attr in self.attribs:
                self.node.set(attr['qname'], attr['val'])
        return self.node

    def add_child(self, parent, child):
        return parent.append(child)