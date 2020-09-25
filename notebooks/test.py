import os
import sys

base_dir        = os.path.dirname(os.getcwd())
sys.path.append(base_dir)

from src.xml.base import DOCX_ELEMENT_FACTORY, NS_DEF_TYPE
from lxml.etree import tounicode


class Section(DOCX_ELEMENT_FACTORY):
    def __init__(self):
        super().__init__()
    
    def get(self):
        '''
            <w:sectPr w:rsidR="00A66D74" w:rsidSect="00034616">
            <w:pgSz w:w="11893" w:h="16840"/>
            <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>
            <w:cols w:space="720"/>
            <w:docGrid w:linePitch="360"/>
            </w:sectPr>
        '''
        root = self.create(NS_DEF_TYPE.w, 'sectPr')
        root.add_attribs(NS_DEF_TYPE.w, 'rsidR', '00A66D74')
        root.add_attribs('rsidSect', '00034616')
        root = super().get()
        
#         self.create(NS_DEF_TYPE.w, 'pgSz')
#         self.add_attribs('w', '11893')
#         self.add_attribs('h', '16840')
#         self.add_child(root, super().get())
        
#         self.create(NS_DEF_TYPE.w, 'pgMar')
#         self.add_attribs('top', '720')
#         self.add_attribs('right', '720')
#         self.add_attribs('bottom', '720')
#         self.add_attribs('left', '720')
#         self.add_attribs('header', '720')
#         self.add_attribs('footer', '720')
#         self.add_attribs('gutter', '0')        
#         self.add_child(root, super().get())

#         self.create(NS_DEF_TYPE.w, 'cols')
#         self.add_attribs('space', '720')
#         self.add_child(root, super().get())

#         self.create_root_node('docGrid')
#         self.add_attribs('linePitch', '360')
#         self.add_child(root, super().get())
        
        return root

node = Section()
node.get()
