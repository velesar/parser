from lxml import etree
from docx import *
import sys, os, fnmatch
import argparse

class myBookTarget(object):

    def __init__(self):
        self.book = {
            'meta':{},
            'chapters':[],
            'imgs':[]
        }

    tag_stack = []
    attrs = {}
    chapters = []


    def meta(self, data):
        self.book['meta'][self.attrs['meta']['name']] = self.attrs['meta']['content']

    def div(self, data):
        self.book['chapters'].append({'number':self.attrs['div'].get('id'),'paragraphs':[]})

    def header(self, data):
        self.book['chapters'][-1]['title'] = data.encode('utf-8')

    def paragraph(self, data):
        paragraph = self.book['chapters'][-1]['paragraphs']
        paragraph.append((data,''))

    def italic(self, data):
        paragraph = self.book['chapters'][-1]['paragraphs']
        paragraph.append((data,'i'))

    tags = {
        'meta':meta,
        'div':div,
        'h2':header,
        'p':paragraph,
        'i':italic
    }

    def start(self, tag, attrib):
        self.tag_stack.append(tag)
        self.attrs[tag] = attrib

    def data(self, data):
        last_tag = self.tag_stack[-1]
        if last_tag in self.tags:
            self.tags[last_tag](self, data)
        
    def end(self, tag):
        self.tag_stack.pop(0)

    def close(self):
        return self.book

# It is better to somehow separate such things from the core logic  
def my_picture(relationshiplist, picname, picdescription, output_dir, pixelwidth=None, pixelheight=None, nochangeaspect=True, nochangearrowheads=True):
    media_dir = join(template_dir, 'word', 'media')
    if not os.path.isdir(media_dir):
        os.mkdir(media_dir)
    shutil.copyfile(picname, join(media_dir, picname.split('/')[-1]))
    #shutil.copyfile(picname, output_dir+picname[picname.rfind('/'):])
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth, pixelheight = Image.open(picname).size[0:2]

    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs
    emuperpixel = 12700
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)

    # Set relationship ID to the first available
    picid = '2'
    picrelid = 'rId'+str(len(relationshiplist)+1)
    relationshiplist.append([
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/'+picname.split('/')[-1]])

    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = makeelement('blipFill', nsprefix='pic')
    blipfill.append(makeelement('blip', nsprefix='a', attrnsprefix='r',
                    attributes={'embed': picrelid}))
    stretch = makeelement('stretch', nsprefix='a')
    stretch.append(makeelement('fillRect', nsprefix='a'))
    blipfill.append(makeelement('srcRect', nsprefix='a'))
    blipfill.append(stretch)

    # 2. The non visual picture properties
    nvpicpr = makeelement('nvPicPr', nsprefix='pic')
    cnvpr = makeelement('cNvPr', nsprefix='pic',
                        attributes={'id': '0', 'name': 'Picture 1', 'descr': picname})
    nvpicpr.append(cnvpr)
    cnvpicpr = makeelement('cNvPicPr', nsprefix='pic')
    cnvpicpr.append(makeelement('picLocks', nsprefix='a',
                    attributes={'noChangeAspect': str(int(nochangeaspect)),
                                'noChangeArrowheads': str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)

    # 3. The Shape properties
    sppr = makeelement('spPr', nsprefix='pic', attributes={'bwMode': 'auto'})
    xfrm = makeelement('xfrm', nsprefix='a')
    xfrm.append(makeelement('off', nsprefix='a', attributes={'x': '0', 'y': '0'}))
    xfrm.append(makeelement('ext', nsprefix='a', attributes={'cx': width, 'cy': height}))
    prstgeom = makeelement('prstGeom', nsprefix='a', attributes={'prst': 'rect'})
    prstgeom.append(makeelement('avLst', nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)

    # Add our 3 parts to the picture element
    pic = makeelement('pic', nsprefix='pic')
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)

    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
    graphicdata = makeelement('graphicData', nsprefix='a',
                              attributes={'uri': 'http://schemas.openxmlforma'
                                                 'ts.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = makeelement('graphic', nsprefix='a')
    graphic.append(graphicdata)

    framelocks = makeelement('graphicFrameLocks', nsprefix='a',
                             attributes={'noChangeAspect': '1'})
    framepr = makeelement('cNvGraphicFramePr', nsprefix='wp')
    framepr.append(framelocks)
    docpr = makeelement('docPr', nsprefix='wp',
                        attributes={'id': picid, 'name': 'Picture 1',
                                    'descr': picdescription})
    effectextent = makeelement('effectExtent', nsprefix='wp',
                               attributes={'l': '25400', 't': '0', 'r': '0',
                                           'b': '0'})
    extent = makeelement('extent', nsprefix='wp',
                         attributes={'cx': width, 'cy': height})
    inline = makeelement('inline', attributes={'distT': "0", 'distB': "0",
                                               'distL': "0", 'distR': "0"},
                         nsprefix='wp')
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
    drawing = makeelement('drawing')
    drawing.append(inline)
    run = makeelement('r')
    run.append(drawing)
    paragraph = makeelement('p')
    paragraph.append(run)
    return relationshiplist, paragraph


class DocCreator(object):
    def doc_create(self):
        self.document = newdocument()
        self.relationships = relationshiplist()
        self.body = self.document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]
    def create_heading(self, data, size = 2):
        self.body.append(heading(data, size))
    def insert_picture(self, name):
    
        picture = my_picture
        self.relationships, picpara = picture(self.relationships, name, '', args.output_dir)
        self.body.append(picpara)
        self.body.append(pagebreak(type='page', orient='portrait'))
    def create_paragraph(self, data):
        self.body.append(paragraph(data))
    def create_list(self, data):
        for item in data:
            self.body.append(paragraph(item))
    def page_break(self):
        self.body.append(pagebreak(type='page', orient='portrait'))
    def doc_save(self, meta, output_dir, title = '', subject = '', creator = '', keywords = []):
        self.title = title
        self.subject = subject
        self.creator = creator
        self.keywords = keywords
        self.coreprops = coreproperties(title=self.title, subject=self.subject, creator=self.creator,keywords=self.keywords)
        self.appprops = appproperties()
        self.contenttypes = contenttypes()
        self.websettings = websettings()
        self.wordrelationships = wordrelationships(self.relationships)
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        savedocx(self.document, self.coreprops, self.appprops, self.contenttypes, self.websettings, self.wordrelationships, '%s/%s %s %s.docx' %(output_dir, meta['booktitle'], meta['author'], meta['year']))

# Hmm... I think this is not very good...
# Will there be a difference if yield replace with return?
# Why did you use 'folder_depth' variable insteed of 'depth'
def find_files(dir, depth, pattern = '*.xml'):
    for i in range(depth):
        for root, dirs, files in os.walk(dir):
            if i <= 2: # what does magic '2' mean?
                for name in files:
                    if fnmatch.fnmatch(name, pattern):
                        filename = os.path.join(root, name)
                        yield filename
            else:
                find_files(root, i)

# Hmm...
# def find_picture(root, depth):
#     picture = ''
#     for files in find_files(root, depth, '*.png'):
#         picture = files
#     if len(picture) == 0:
#         for files in find_files(root, depth, '*.jpg'):
#             picture = files
#     return picture

def lookup_for_imgs(book, file_dir):
    # import pudb
    # pudb.set_trace()
    for m_file in os.listdir(file_dir):
        if fnmatch.fnmatch(m_file, '*.png') or fnmatch.fnmatch(m_file, '*.jpg'):
            book['imgs'].append(os.path.join(file_dir,m_file))
    return book


class TemplateCreator(object):
    def read(self, name):
        self.f = open(name, 'r')
    def make(self, result, output_dir):
        self.doc = DocCreator()
        self.doc.doc_create()
        # pict = find_picture()
        # if pict and pict[:pict.rfind('/')] == file_xml[:file_xml.rfind('/')]:
        #     self.doc.insert_picture(pict)
        for pict in result['imgs']:
            self.doc.insert_picture(pict)
        self.doc.create_heading(result['meta']['booktitle'])
        self.doc.create_heading(result['meta']['author'])
        self.doc.create_paragraph('%s (%s)' %(result['meta']['year'],result['meta']['firstpub']))
        self.doc.page_break()
        self.doc.create_heading(result['meta']['publisher'])
        self.doc.create_heading(result['meta']['address'])
        self.doc.page_break()
        for i in range(len(result['chapters'])):
            for item in result['chapters'][i]:
                if item == 'title':
                    self.doc.create_paragraph(result['chapters'][i][item].decode('utf-8'))
        
        for i in range(len(result['chapters'])):
            for item in result['chapters'][i]:
                if item == 'title':
                    self.doc.page_break()
                    self.doc.create_heading(result['chapters'][i][item].decode('utf-8'))
                if item == 'paragraphs':
                    self.doc.create_paragraph(result['chapters'][i][item])

        self.doc.page_break()
        self.doc.doc_save(result['meta'], output_dir)      

def parse(rootdir, folder_depth, max_count, output_dir):

    filexml = [] # >>- Collection is better to call using the plural
    results = []

    # >>- BTW. If you need to process 5 files from 50k xmls.
    # >>       How much time the script will do unnecessary work??? ))))
    # >>       It is about -m options.
    for files in find_files(rootdir, folder_depth):
        if files not in filexml:
            filexml.append(files)    

    temp = TemplateCreator()

    
    # for i in range(len(filexml)):
    # >> How can there be another type of variable ?
    #     if int(args.m) != 0:
    #         if count <= args.m:
    #             results.append(etree.parse(filexml[i], parser))
    #             temp.make(results[i], filexml[i])
    #             count += 1 
    #     else: 
    #         results.append(etree.parse(filexml[i], parser))
    #         temp.make(results[i], filexml[i])


    if max_count != 0:
        filexml = filexml[0:max_count]

    for xml_file in filexml:
        book = etree.parse(xml_file, etree.XMLParser(target = myBookTarget()))
        book = lookup_for_imgs(book, os.path.dirname(xml_file))
        temp.make(book, output_dir)
        
if __name__ == '__main__':
    
    # I think it will be better
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument("rootdir", type = str)
    arg_parser.add_argument("-d", nargs='?', type = int, default = 2)
    arg_parser.add_argument("-m", nargs='?', type = int, default = 0)
    arg_parser.add_argument("output_dir", nargs='?', type = str, default ='output')
    # and what about <word_template.docx> - argument ?
    args = arg_parser.parse_args()

    parse(args.rootdir, args.d, args.m, args.output_dir)
